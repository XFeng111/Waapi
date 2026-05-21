class WwiseBase:
    def __init__(self, client:object):
        # 传入外部的client,client = WaapiClient()
        self.client = client

class Core_object(WwiseBase):
    def getChild_SoundId(self, object_id, depth:int=0):
        child_list = []
        max_depth = 100
        if depth > max_depth:
            print(f"递归深度超过限制，终止：object_id={object_id}\n")
            exit() # 终止脚本

        res_children = self.object_get(object_id, opt=["children.id", "children.type"])['return']
        for child in res_children:
            child_id_list = child['children.id']
            child_type_list = child['children.type']
            for child_id, child_type in zip(child_id_list, child_type_list):
                if child_type == "Sound":
                    child_list.append(child_id)
                else:
                    child_list += self.getChild_SoundId(child_id, depth+1)
            
            if child_id_list == [] and child_type_list == []:
                print(f"未找到子对象：object_id={object_id}")
                exit() # 终止脚本

        return child_list
    # res_child = getChild_SoundId("{874CC972-8752-4D28-9563-0ABFCEFA1DCB}")
    # pprint(res_child)

    def audio_import(self, originalsSubFolder, audioFile, objectPath, objectType, opt:list, importOperation:str="useExisting", lang_marker:str="SFX"):
        args = {
            # useExisting ：使用现有对象（如有），更新给定属性；否则，创建新的对象。该项为默认值。
            # replaceExisting ：创建新的对象；若存在同名的现有对象，则将现有对象销毁。
            # createNew ：创建新的对象；在可能的情况下赋予对象以所需名称，否则使用新的唯一名称。

            "importOperation": importOperation, 
            "default": {
                "importLanguage": lang_marker
            },
            "imports": [
                {
                    "originalsSubFolder":originalsSubFolder,
                    "audioFile": audioFile,
                    "objectPath": objectPath,
                    "objectType":objectType
                }
            ]
        }
        options = {
            "return": opt
        }
        return self.client.call("ak.wwise.core.audio.import", args, options=options)

    def setProperty(self, object_id, property, value):
        args ={
                "object": object_id,
                "property": property,
                "value": value
                }
        return self.client.call("ak.wwise.core.object.setProperty",args)
    
    def setReference(self, object_id, reference, value):
        args ={
                "object": object_id,
                "reference": reference,
                "value": value
                }
        return self.client.call("ak.wwise.core.object.setReference",args)
    
    def switchContainer_addAssignment(self, child, stateOrSwitch):
        args ={
                "child": child,
                "stateOrSwitch": stateOrSwitch
                }
        return self.client.call("ak.wwise.core.switchContainer.addAssignment",args)
    
    def setName(self, object_id, new_name):
        args = {
            "object": object_id,
            "value": new_name
        }
        return self.client.call("ak.wwise.core.object.setName", args)

    def pasteProperties(self, source_id, targets_id, pasteMode:str="replaceEntire"):
        args = {
            "source": source_id,
            "targets": [targets_id],
            "pasteMode": pasteMode
            # inclusion:list = [], 所要包含的属性、引用和列表
            # exclusion :list = [], 所要排除的属性、引用和列表
        }
        
        return self.client.call("ak.wwise.core.object.pasteProperties", args)

    def object_create(self, parent, name, obj_type, onNameConflict:str="merge"):
        args = {
            "parent":parent,
            "name":name,
            "type":obj_type,
            "onNameConflict": onNameConflict
        }
        return self.client.call("ak.wwise.core.object.create", args)

    def object_delete(self, object):
        args = {
            "object": object
        }
        return self.client.call("ak.wwise.core.object.delete", args)

    def object_get(self, object_id, opt:list):
        args = {
            "waql": f"$\"{object_id}\""
        }
        options = {
            "return": opt
        }
        return self.client.call("ak.wwise.core.object.get", args, options=options)
        # ['return'][0]['...']['...'] 取值
    
    def object_move(self, object_id, parent, onNameConflict:str="fail"):
        args = {
            "object":object_id,
            "parent":parent,
            "onNameConflict": onNameConflict # rename\replace\fail（默认）
        }
        return self.client.call("ak.wwise.core.object.move", args)
    
    def play_event_create(self, 
                          event_name, 
                          target_id, 
                          parent_path:str="\\Events", 
                          parent_type:str="WorkUnit", 
                          parent_name:str="Default Work Unit", 
                          onNameConflict:str="merge"):
        args = {
            "parent": parent_path,
            "type": parent_type,
            "name": parent_name,
            "onNameConflict": onNameConflict,
            "children": [
                {
                    "type": "Event",
                    "name": f"Play_{event_name}",
                    "children":[
                        {
                            "name": "",
                            "type": "Action",
                            "@ActionType": 1,
                            "@Target": target_id
                        }
                    ]
                }
            ]
        }
        print(f"✅ 创建事件 Play_{event_name}, 路径：{parent_path}\\{parent_name}\\Play_{event_name}")
        return self.client.call("ak.wwise.core.object.create", args)
    
    def stop_event_create(self, 
                          event_name, 
                          target_id, 
                          parent_path:str="\\Events", 
                          parent_type:str="WorkUnit", 
                          parent_name:str="Default Work Unit", 
                          onNameConflict:str="merge"):
        args = {
            "parent": parent_path,
            "type": parent_type,
            "name": parent_name,
            "onNameConflict": onNameConflict,
            "children": [
                {
                    "type": "Event",
                    "name": f"Stop_{event_name}",
                    "children":[
                        {
                            "name": "",
                            "type": "Action",
                            "@ActionType": 2,
                            "@Target": target_id,
                            "@Scope": 1,
                            "@FadeTime": 0.8
                        }
                    ]
                }
            ]
        }
        print(f"✅ 创建事件 Stop_{event_name}, 路径：{parent_path}\\{parent_name}\\Stop_{event_name}")
        return self.client.call("ak.wwise.core.object.create", args)
    
    def event_creat_FromOld(self, object_name, object_id, ActionType:int=1): # 1-Play, 2-Stop
        object_parent_id = self.object_get(object_id, ["parent.id"]) ['return'][0]['parent.id']
        parent_children = self.object_get(object_parent_id, ["children.id"]) ['return'][0]['children.id']

        for child_id in parent_children:
            child_refer = self.object_get(child_id, ["referencesTo"]) ['return'][0]
            e_parent_path = ""
            if child_refer != {}:
                refer_id = child_refer['referencesTo'][0]['id']
                event_info = self.object_get(refer_id, ["parent.parent.path", "parent.parent.type"]) ['return'][0]
                e_parent_path = event_info['parent.parent.path']
                e_parent_type = event_info['parent.parent.type']
                print(f"找到已有资源的Event创建路径：{e_parent_path}，类型：{e_parent_type}")
                break

        if e_parent_path != "" and e_parent_path != "\\Events":
            e_parent_path_parts = [part for part in e_parent_path.split("\\") if part]
            e_parent_name = e_parent_path_parts[-1]
            e_parent_Cr_path = ""
            for p in e_parent_path_parts[:-1]: # 从开头 取到 倒数第二个 结束
                e_parent_Cr_path = e_parent_Cr_path+"\\"+p

            if ActionType == 1:
                self.play_event_create(object_name, object_id, e_parent_Cr_path, e_parent_type, e_parent_name, "merge")
                return "play event"
            
            if ActionType == 2:
                self.stop_event_create(object_name, object_id, e_parent_Cr_path, e_parent_type, e_parent_name, "merge")
                return "stop event"
            
        elif e_parent_path == "\\Events": # 默认Event路径创建
            if ActionType == 1:
                self.play_event_create(object_name, object_id)
                return "play event"
            
            if ActionType == 2:
                self.stop_event_create(object_name, object_id)
                return "stop event"

        else:
            print(f"error: 没有找到Event路径，请检查 {object_name} 同层级的已有资源是否有创建Event")
            return "error"
    # c_obj.event_creat_FromOld(object_name, object_id, 1) # play event
    # c_obj.event_creat_FromOld(object_name, object_id, 2) # stop event

    def event_creat_FromActorPath(self, object_name, object_id, ActionType:int=1): # 1-Play, 2-Stop
        object_parent_info = self.object_get(object_id, ["parent.path", "parent.id"])
        object_parent_path = object_parent_info ['return'][0]['parent.path']
        object_parent_id = object_parent_info ['return'][0]['parent.id']

        parent_children = self.object_get(object_parent_id, ["children.id"]) ['return'][0]['children.id']
        child_id = parent_children[0]

        child_refer = self.object_get(child_id, ["referencesTo"]) ['return'][0]
        o_parent_path_parts = [part for part in object_parent_path.split("\\") if part]

        if child_refer == {}:

            e_parent_Cr_path = "\\Events"
            e_parent_name = o_parent_path_parts[1]

            # 创建对应 Work Unit，注意WorkUnit 会打断前面的 undo 组
            self.object_create("\\Events", e_parent_name, "WorkUnit", "merge")

            # 创建对应的文件夹结构
            if len(o_parent_path_parts)-1 >= 2:
                for i in range(1, len(o_parent_path_parts)-1): # 从第二个取到结束
                    e_parent_Cr_path = e_parent_Cr_path+"\\"+ o_parent_path_parts[i]
                    e_parent_name = o_parent_path_parts[i+1] # 获取父级文件夹名称
                    self.object_create(e_parent_Cr_path, e_parent_name, "Folder", "merge")
            
            print("未找到已有资源的Event结构，已按照ActorMixer创建Event结构")
            
            if ActionType == 1:
                self.play_event_create(object_name, object_id, e_parent_Cr_path, "Folder", e_parent_name, "merge")
                return "play event"
            
            if ActionType == 2:
                self.stop_event_create(object_name, object_id, e_parent_Cr_path, "Folder", e_parent_name, "merge")
                return "stop event"

            print("已按照ActorMixer结构创建Event")

        else:
            if ActionType == 1:
                self.event_creat_FromOld(object_name, object_id, 1) # play event
                return "play event"
        
            if ActionType == 2:
                self.event_creat_FromOld(object_name, object_id, 2) # stop event
                return "stop event"
    # c_obj.event_creat_FromActorPath(object_name, object_id, 1) # play event
    # c_obj.event_creat_FromActorPath(object_name, object_id, 2) # stop event

    def sourceControl_add(self, file):
        args = {
            "files": [file]
        }
        return self.client.call("ak.wwise.core.sourceControl.add", args)
    
    def sourceControl_delete(self, file):
        args = {
            "files": [file]
        }
        return self.client.call("ak.wwise.core.sourceControl.delete", args)


class Core_undo(WwiseBase):
    def undo_beginGroup(self):
        return self.client.call("ak.wwise.core.undo.beginGroup")
    
    def undo_endGroup(self, displayName:str):
        return self.client.call("ak.wwise.core.undo.endGroup", {"displayName": displayName})
        # displayName 在历史记录中针对此 Undo Group 显示的名称
        # client.call("ak.wwise.core.undo.endGroup", {"displayName": "T_Event_Creat_FromActorMixer"})


class Ui(WwiseBase):
    def getSelectedObjects(self, opt:list):
        options= {
            "return":opt
        }
        return self.client.call("ak.wwise.ui.getSelectedObjects",options=options)
        # ["objects"][0]["..."]["..."] 取值
