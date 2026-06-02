from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import msvcrt
from WAAPI_Functions import Core_object, Core_undo, Ui
from T_Rename_WavAndEvent_FromSound import Rename_FromSound

def T_Sound_ToRandomContainer():
    # 初始化类，传入client
    c_obj = Core_object(client)
    ui = Ui(client)
    # 获取选中的对象
    opt = ["id","name", "type", "parent.path", "referencesTo.id"]
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
        obj_name = obj['name']
        obj_parent_path = obj['parent.path']
        obj_ref_list = obj['referencesTo.id']

        c_obj.setName(obj_id, obj_name + "_01") # 重命名Sound对象
        c_obj.object_create(obj_parent_path, obj_name, "RandomSequenceContainer") # 创建随机容器
        c_obj.object_move(obj_id, obj_parent_path + "/" + obj_name) # 将Sound对象移动到随机容器下

        for ref in obj_ref_list:
            c_obj.setReference(ref, "Target", obj_parent_path + "/" + obj_name)
            print(f"✅ 已更新 {obj_name} 引用")

        Rename_FromSound(c_obj, obj_id) # 调用重命名函数，传入Sound对象的id
        
        print(f"✅ 已完成 {obj_name} Sound对象转换为随机容器")


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

        T_Sound_ToRandomContainer()

        # 结束某个 Undo Group
        c_undo.undo_endGroup("T_Sound_ToRandomContainer")

        # 断开连接
        client.disconnect() # 必须手动关闭连接，否则文件会被占用
        print("✅ 已断开 WAAPI 连接")

    finally:
        # input("\n按 Enter 键退出...") # 等待用户按下 Enter 键,macOS/Linux使用
        print("\n按任意键退出...")
        msvcrt.getch() # 等待用户按下任意键,Windows使用