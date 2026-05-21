from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import msvcrt
from WAAPI_Functions import Core_object, Core_undo, Ui

def Source_SetLoop(c_obj:Core_object, obj_id):

    obj_set = c_obj.object_get(obj_id, ["id","name","activeSource","duration", "IsLoopingEnabled"])['return'][0]
    obj_set_id = obj_set["id"]
    obj_set_name = obj_set["name"]
    source_id = obj_set["activeSource"]['id']
    source_name = obj_set["activeSource"]['name']
    source_duration = obj_set['duration']['min']
    
    if obj_set['IsLoopingEnabled'] == True:
        # 如果已经开启Loop，Loop状态下的duration为浮点数，关闭Loop重新取关闭状态下的duration
        c_obj.setProperty(obj_id, "IsLoopingEnabled", False)
        source_duration_False = c_obj.object_get(obj_id, ["duration"])['return'][0]['duration']['min']
        source_duration = source_duration_False
        # pprint(source_duration_False)
    
    args_list = [
        (obj_set_id, "IsLoopingEnabled", True),
        (source_id, "OverrideWavLoop", True),
        (source_id, "FadeInDuration", 0.3),
        (source_id, "LoopBegin", 0.2*source_duration),
        (source_id, "LoopEnd", 0.9*source_duration),
        (source_id, "CrossfadeDuration", 0.06*source_duration*1000)
    ]

    for object, property, value in args_list:
        c_obj.setProperty(object, property, value)
        print(f" ✅ {obj_set_name} 的 {property} 设置为 {value}")

    print(f"✅ 成功开启【{obj_set_name}】的Loop \n 【{source_name}】的OverrideWavLoop\n")

def T_ActiveSource_Loop():
    # 初始化类，传入client
    c_obj = Core_object(client)
    ui = Ui(client)
    # 获取选中的对象
    opt = ["id","name", "type"]
    objects = ui.getSelectedObjects(opt)
    obj_list = objects.get("objects")

    # debug数据
    # pprint(obj_list)

    # 遍历所有选中的对象，逐个设置
    for obj in obj_list:
        if obj['type'] != 'Sound':
            print(f"❌ {obj['name']} 不是Sound,请在Wwise中选中Sound对象后重新运行脚本")
            # return exit() # 退出脚本
            return None # 给函数返回None值，终止当前所在的函数执行，并返回到调用该函数的地方继续执行

        obj_id = obj['id']
        Source_SetLoop(c_obj, obj_id)

if __name__ == "__main__":
    try:
        client = WaapiClient()
        print("✅ 已连接 Wwise WAAPI")
    except CannotConnectToWaapiException as e:
        print(f"❌ 无法连接 WAAPI:{e},请打开 Wwise 并确保 WAAPI 已启用")
    else:
        # 开始某个 Undo Group
        c_undo = Core_undo(client)
        c_undo.undo_beginGroup()

        T_ActiveSource_Loop()

        # 结束某个 Undo Group
        c_undo.undo_endGroup("T_ActiveSource_Loop")

        # 断开连接
        client.disconnect() # 必须手动关闭连接，否则文件会被占用
        print("✅ 已断开 WAAPI 连接")

    finally:
        # input("\n按 Enter 键退出...") # 等待用户按下 Enter 键,macOS/Linux使用
        print("\n按任意键退出...")
        msvcrt.getch() # 等待用户按下任意键,Windows使用
