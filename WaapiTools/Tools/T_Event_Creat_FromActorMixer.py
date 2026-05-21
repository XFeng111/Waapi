from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import msvcrt
from WAAPI_Functions import Core_object, Core_undo, Ui
from T_ActiveSource_Loop import Source_SetLoop

def T_Event_Creat_FromActorMixer():
    index = input("输入y —— 从已有资源的Event结构创建，若无Event结构则按照ActorMixer创建\n输入n —— 创建至默认路径Events\n等待输入ing...")

    # 初始化类，传入client
    c_obj = Core_object(client)
    ui = Ui(client)

    # 获取选中的对象
    opt = ["name", "type", "children"]
    objects = ui.getSelectedObjects(opt)["objects"]
    if objects == []:
        print("❌ 未选中任何对象,请选择 ActorMixer 后重新运行脚本")
        return None # 给函数返回None值，终止当前所在的函数执行，并返回到调用该函数的地方继续执行
        # return exit() # 如果没有选中任何对象，则退出程序

    # pprint(objects) # Debug打印对象信息
    for obj in objects:
        obj_name = obj["name"]
        obj_type = obj["type"]
        obj_child = obj["children"]

        if obj_type != "ActorMixer":
            print(f"{obj_name}不是ActorMixer，请选择 ActorMixer 后重新运行脚本")
        else:
            for o in obj_child:
                o_name = o["name"]
                o_id = o["id"]
                o_refer = c_obj.object_get(o_id, ["referencesTo"])
                o_type = c_obj.object_get(o_id, ["type"])["return"][0]["type"]

                if o_type == "ActorMixer" or o_type == "Folder":
                    print(f"{o_name}是ActorMixer或Folder，已跳过新建Event\n")
                else:
                    if o_refer == {'return': [{}]}:
                        if index.lower() == "n":
                            # 默认路径创建事件
                            c_obj.play_event_create(o_name, o_id)
                            # 若_loop或_lp结尾，则继续创建停止事件，并为Source设置Loop
                            if o_name.lower().endswith(("_loop", "_lp")):
                                # 创建停止事件
                                c_obj.stop_event_create(o_name, o_id)
                                # 若为Sound，为Source设置Loop
                                if o_type == "Sound":
                                    Source_SetLoop(c_obj, o_id)
                        elif index.lower() == "y":

                            client.call("ak.wwise.core.undo.endGroup", {"displayName": "Before Creat Work Unit"})
                            # ------------------------------------------------
                            c_obj.event_creat_FromActorPath(o_name, o_id, 1) # WorkUnit 会打断前面的 undo 组，补一组防报错
                            # ------------------------------------------------
                            client.call("ak.wwise.core.undo.beginGroup")

                            if o_name.lower().endswith(("_loop", "_lp")):
                                c_obj.event_creat_FromActorPath(o_name, o_id, 2)
                                if o_type == "Sound":
                                    Source_SetLoop(c_obj, o_id)

                        else:
                            print("❌ 输入有误，请重新运行脚本")
                            return

                    else:
                        o_refer_count = len(o_refer["return"][0]["referencesTo"])
                        print(f"{o_name}已被{o_refer_count}个对象引用，已跳过新建Event\n")

if __name__ == "__main__":
    try:
        client = WaapiClient()
        print("✅ 已连接 Wwise WAAPI")
    except CannotConnectToWaapiException as e:
        print(f"❌ 无法连接 WAAPI:{e},请打开 Wwise 并确保 WAAPI 已启用")
    else:
        # 开始某个 Undo Group
        c_undo = Core_undo(client)

        c_undo.undo_beginGroup() # 涉及创建 Work Unit 会打断undo组
        T_Event_Creat_FromActorMixer()

        # 结束某个 Undo Group
        c_undo.undo_endGroup("T_Event_Creat_FromActorMixer")

        # 断开连接
        client.disconnect() # 必须手动关闭连接，否则文件会被占用
        print("✅ 已断开 WAAPI 连接")

    finally:
        # input("\n按 Enter 键退出...") # 等待用户按下 Enter 键,macOS/Linux使用
        print("\n按任意键退出...")
        msvcrt.getch() # 等待用户按下任意键,Windows使用
    
