from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import msvcrt
from WAAPI_Functions import Core_object, Core_undo, Ui

# name_mode=1 时，加 _Player尾缀，Player命名为 "SwitchName _Player"，如 "VO _Player"
# name_mode=0 时，不加 _Player尾缀，Player命名为 "SwitchName"，如 "VO"
def T_Vo_SetPlayerSwitch_FromActorMixer(switch_group_path, name_mode=1):
    # 初始化类，传入client
    c_obj = Core_object(client)
    ui = Ui(client)
    # 获取选中的对象
    opt = ["path","name", "type"]
    objects = ui.getSelectedObjects(opt)
    obj_list = objects.get("objects")

    # debug数据
    # pprint(obj_list) 

    # 遍历所有选中的对象，逐个设置
    for obj in obj_list:
        if obj['type'] != 'ActorMixer':
            print(f"❌ {obj['name']} 不是Actor Mixer,请在Wwise中选中Actor Mixer对象后重新运行脚本")
            # return exit() # 退出脚本
            return None # 给函数返回None值，终止当前所在的函数执行，并返回到调用该函数的地方继续执行

        obj_path = obj['path']
        child_path_list = []
        child_name_list = []
        switch_name_list = []

        res_children = c_obj.object_get(obj_path, opt=["children.path", "children.type", "children.name"])['return']
        for child in res_children:
            child_path_list = child['children.path']
            child_type_list = child['children.type']
            child_name_list = child['children.name']
            for child_path, child_type, child_name in zip(child_path_list, child_type_list, child_name_list):
                if child_type == "Sound":
                    if child_name.lower().endswith(("female", "male")):
                        print(f"找到{child_name}")
                        switch_name = child_name.rsplit("_", 1)[0] # 去掉Female\Male后缀
                        if switch_name not in switch_name_list:
                            switch_name_list.append(switch_name)

                else:
                    print(f"{child_name} ({child_path}) 不是Sound对象，已跳过")
            
            if child_path_list == [] and child_type_list == []:
                print(f"未找到子对象：{obj_path}")
                exit() # 终止脚本

        # Player创建Switch Container对象
        if switch_name_list == []:
            print(f"❌ 在 {obj_path} 中未找到Female\Male")
            exit() # 终止脚本
        else:
            for switch_name in switch_name_list:
                if name_mode == 1:
                    player_name = switch_name + "_Player"
                elif name_mode == 0:
                    player_name = switch_name
                
                player_path = obj_path + "\\" + player_name
                f_path = obj_path + "\\" + switch_name + "_Female"
                m_path = obj_path + "\\" + switch_name + "_Male"

                player_F_path = obj_path + "\\" + player_name + "\\" + switch_name + "_Female"
                player_M_path = obj_path + "\\" + player_name + "\\" + switch_name + "_Male"

                c_obj.object_create(obj_path, player_name, "Switch Container", "merge") # Player创建Switch Container对象
                c_obj.object_move(f_path, player_path, onNameConflict="fail") # Female移动到player_path下
                c_obj.object_move(m_path, player_path, onNameConflict="fail") # Male移动到player_path下

                c_obj.setReference(player_path, "SwitchGroupOrStateGroup", switch_group_path) # Player设置Switch Group属性
                c_obj.setReference(player_path, "DefaultSwitchOrState", switch_group_path+"\\Male") # Player设置DefaultSwitchOrState属性

                c_obj.switchContainer_addAssignment(player_F_path, switch_group_path+"\\Female")
                c_obj.switchContainer_addAssignment(player_M_path, switch_group_path+"\\Male")

                if "\\States" in switch_group_path:
                    c_obj.switchContainer_addAssignment(player_M_path, switch_group_path+"\\None")

                print(f"✅ 已创建 {player_name} 并设置Switch Group属性")
        
        print(f"{obj_path} 已完成设置\n")

if __name__ == "__main__":
    try:
        client = WaapiClient()
        print("✅ 已连接 Wwise WAAPI\n")
    except CannotConnectToWaapiException as e:
        print(f"❌ 无法连接 WAAPI:{e},请打开 Wwise 并确保 WAAPI 已启用")
    else:
        # 开始某个 Undo Group
        c_undo = Core_undo(client)
        c_undo.undo_beginGroup()

        index = input("选择Switch Group路径\n输入1 —— \\States\\System\\Player_Gender\n输入2 —— \\Switches\\Default Work Unit\\Gender_World\n等待输入ing...")

        if index == "1":
            switch_group_path = "\\States\\System\\Player_Gender"
        elif index == "2":
            switch_group_path = "\\Switches\\Default Work Unit\\Gender_World"
        else:
            print("❌ 无效输入")
            exit()

        T_Vo_SetPlayerSwitch_FromActorMixer(switch_group_path, 1) # name_mode=1 时，Player命名为 "SwitchName _Player"，如 "VO _Player"

        # 结束某个 Undo Group
        c_undo.undo_endGroup("T_Vo_SetPlayerSwitch_FromActorMixer")

        # 断开连接
        client.disconnect() # 必须手动关闭连接，否则文件会被占用
        print("✅ 已断开 WAAPI 连接")

    finally:
        # input("\n按 Enter 键退出...") # 等待用户按下 Enter 键,macOS/Linux使用
        print("\n按任意键退出...")
        msvcrt.getch() # 等待用户按下任意键,Windows使用