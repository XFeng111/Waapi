from waapi import WaapiClient, CannotConnectToWaapiException
from pprint import pprint
import msvcrt
from WAAPI_Functions import Core_object, Core_undo, Ui
import shutil
import os


def copy_wav_file(source_wav_path, target_dir, new_filename):
    """
    将源WAV文件复制到指定路径，并自定义新文件名
    
    参数:
        source_wav_path (str): 源WAV文件的完整路径（如 E:\\test.wav）
        target_dir (str): 目标文件夹路径（如 D:\\audio\\）
        new_filename (str): 新文件的名称（无需后缀，函数会自动添加.wav）
    
    返回:
        str: 复制后的文件完整路径；若失败返回空字符串
    
    异常处理:
        捕获文件不存在、路径权限、文件夹不存在等常见错误
    """
    try:
        # 1. 校验源文件是否存在且是.wav文件
        if not os.path.exists(source_wav_path):
            print(f"错误：源文件 {source_wav_path} 不存在")
            return ""
        if not source_wav_path.lower().endswith(".wav"):
            print(f"错误：{source_wav_path} 不是WAV文件")
            return ""
        
        # 2. 确保目标文件夹存在，不存在则创建
        os.makedirs(target_dir, exist_ok=True)
        
        # 3. 拼接目标文件完整路径（自动添加.wav后缀）
        target_wav_path = os.path.join(target_dir, f"{new_filename}.wav")
        
        # 4. 复制文件（shutil.copy保留文件元数据，copy2更完整）
        shutil.copy2(source_wav_path, target_wav_path)
        
        print(f"文件复制成功，新命名wav路径：{target_wav_path}")
        return target_wav_path
    
    except PermissionError:
        print(f"错误：没有权限访问 {target_dir} 或源文件")
        return ""
    except Exception as e:
        print(f"复制失败：未知错误 - {str(e)}")
        return ""
    
def get_originalsSubFolder(file_path, marker="SFX"):
    markers = [marker, "CN", "EN"]
    wav_marker = ".wav"
    # 先检查路径中是否同时包含 marker\\和.wav
    for new_marker in markers:
        if (new_marker+"\\") in file_path and wav_marker in file_path:
            dir_path = os.path.dirname(file_path)
            # 找到 new_marker 的结束位置
            sfx_end_index = file_path.find(new_marker) + len(new_marker)
            # 提取中间的字段（去除可能的路径分隔符）
            target_str = dir_path[sfx_end_index:].strip("\\")
            
            return target_str, new_marker
    
    # 如果所有 marker 都不在 file_path 中，则返回空
    return "", ""

def unusedSources_delete(c_obj:Core_object, sound_id, obj_name):
    res_children = c_obj.object_get(sound_id, opt=["children"])["return"][0]['children']
    for child in res_children:
        c_id = child["id"]
        c_name = child["name"]
        if c_name != obj_name:
            c_obj.object_delete(c_id)
            print(f"AudioFile已移除未使用的{c_name}.wav")

def event_targets_pasteProperties(c_obj:Core_object, event_source_id, event_target_id, pasteMode="replaceEntire"):
    tar_source = c_obj.object_get(event_source_id, ["children.id"])["return"][0]['children.id']
    tar_target = c_obj.object_get(event_target_id, ["children.id"])["return"][0]['children.id']
    for s_id, t_id in zip(tar_source, tar_target):
        c_obj.pasteProperties(s_id, t_id, pasteMode)
    # print(f"Event的Target属性已复制")


def T_Rename_WavAndEvent_FromSound(): # 主函数
    # 初始化类，传入client
    c_obj = Core_object(client)
    ui = Ui(client)
    objects = ui.getSelectedObjects(["id", "type"])
    obj_list = objects.get("objects")
    # debug数据
    # pprint(obj_list)
    for obj in obj_list:
        obj_id = obj["id"]
        obj_type = obj["type"]
        if obj_type == "Sound":
            Rename_FromSound(c_obj, obj_id)
        else:
            sound_list = c_obj.getChild_SoundId(obj_id)
            for sound_id in sound_list:
                Rename_FromSound(c_obj, sound_id)
    print("\n--重命名wav,event执行完成--\n")

def Rename_Event(c_obj:Core_object, event_id, event_name, new_name):
    if event_name[:4] == "Play":
        c_obj.setName(event_id, "Play_"+new_name)
        print(f"{event_name}: Event已重命名 --> Play_{new_name}\n")
    
    if event_name[:4] == "Stop":
        c_obj.setName(event_id, "Stop_"+new_name)
        print(f"{event_name}: Event已重命名 --> Stop_{new_name}\n")


def Rename_Wav(c_obj:Core_object, obj_originalWavFilePath, new_filename, sound_path):
    originalsSubFolder, lang_marker = get_originalsSubFolder(obj_originalWavFilePath)
    target_dir = os.path.dirname(obj_originalWavFilePath)
    audioFile = os.path.join(target_dir, f"{new_filename}.wav")
    objectPath = sound_path
    objectType = "Sound" 
    opt = ["id", "name", "path", "sound:originalWavFilePath"]

    file_copy = copy_wav_file(obj_originalWavFilePath, target_dir, new_filename)

    res_import = c_obj.audio_import(originalsSubFolder, audioFile, objectPath, objectType, opt, "useExisting", lang_marker)
    pprint(res_import)
    # print(f"Wav已重命名\n")


def Rename_FromSound(c_obj:Core_object, object_id): # 以Sound类型为基准重命名wav，event，删除未使用的wav，并添加版本控制
    # 获取选中的对象
    opt = ["id", "name", "id", "path", "sound:originalWavFilePath", "referencesTo.id"]
    objects = c_obj.object_get(object_id, opt)
    obj_list = objects.get("return")

    for obj in obj_list:
        obj_name = obj["name"] 
        obj_id = obj["id"]
        obj_path = obj["path"] 
        obj_originalWavFilePath = obj.get("sound:originalWavFilePath", [])
        obj_referencesTo_id = obj.get("referencesTo.id", [])

        if obj_originalWavFilePath!= []:
            target_dir = os.path.dirname(obj_originalWavFilePath)
            audioFile = os.path.join(target_dir, f"{obj_name}.wav")
            filename = os.path.splitext(os.path.basename(obj_originalWavFilePath))[0]

            if filename != obj_name:
                Rename_Wav(c_obj, obj_originalWavFilePath, obj_name, obj_path)
                unusedSources_delete(c_obj, obj_id, obj_name)
                print(f"{obj_name}: Wav已重命名")

                try:
                    res_add = c_obj.sourceControl_add(audioFile)
                    # pprint(res_add)
                    res_add = c_obj.sourceControl_delete(obj_originalWavFilePath)
                    # pprint(res_add)
                    print(f"{obj_name}: sourceControl执行成功")

                except Exception as e:
                    print(f"{obj_name}: sourceControl执行失败：{e}\n")
            else:
                print(f"{obj_name}: wav名称一致，无需修改，已跳过执行")
        else:
            print(f"{obj_name}: 未找到wav文件，已跳过执行")

        if obj_referencesTo_id != []:
            for refer_id in obj_referencesTo_id:
                event = c_obj.object_get(refer_id, ["parent.name", "parent.id"])["return"][0]
                event_name = event["parent.name"]
                event_id = event["parent.id"]
                if obj_name != event_name[5:]:
                    Rename_Event(c_obj, event_id, event_name, obj_name)
                    # print(f"{event_name}: Event已重命名\n")
                else:
                    print(f"{event_name}: event名称一致，无需修改，已跳过执行\n")
        else:
            print(f"{obj_name}: 未找到event，已跳过执行\n")

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

        T_Rename_WavAndEvent_FromSound() # 调用主函数

        # 结束某个 Undo Group
        c_undo.undo_endGroup("T_Rename_WavAndEvent_FromSound")

        # 断开连接
        client.disconnect() # 必须手动关闭连接，否则文件会被占用
        print("✅ 已断开 WAAPI 连接")

    finally:
        # input("\n按 Enter 键退出...") # 等待用户按下 Enter 键,macOS/Linux使用
        print("\n按任意键退出...")
        msvcrt.getch() # 等待用户按下任意键,Windows使用