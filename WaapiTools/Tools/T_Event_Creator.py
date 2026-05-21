"""Event Creator - WAAPI Event 创建工具

为 ActorMixer 下未被引用的子对象创建 Play Event，
可选创建 Stop Event（按名称规则或 All Stop）。
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
import sys
import os

# 添加项目根目录到 sys.path
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.abspath(os.path.join(_SCRIPT_DIR, "..", ".."))
sys.path.insert(0, _PROJECT_ROOT)

from waapi import WaapiClient


class EventCreatorApp:
    """Event 创建工具 UI 应用"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Event Creator - WAAPI")
        self.root.resizable(True, True)

        # 数据存储
        self.actor_mixer_entries = []
        self.event_entries = []
        self.stop_name_entries = []
        self.all_stop_var = None

        self._build_ui()
        self._init_rows()

    def _build_ui(self):
        """构建 UI 界面"""
        # ---- ActorMixer / Event 路径区 ----
        path_frame = ttk.LabelFrame(self.root, text="路径映射")
        path_frame.pack(fill="x", padx=8, pady=4)

        # 表头
        header = ttk.Frame(path_frame)
        header.pack(fill="x", padx=4, pady=2)
        ttk.Label(header, text="ActorMixer路径").pack(side="left", padx=2)
        # 用不可见占位填充 ActorMixer 输入框+按钮的宽度
        ttk.Label(header, text=" ", width=27).pack(side="left", padx=2)
        ttk.Label(header, text="Event路径").pack(side="left", padx=2)

        # 行容器
        self.row_container = ttk.Frame(path_frame)
        self.row_container.pack(fill="x", padx=4, pady=2)

        # 增删按钮
        btn_frame = ttk.Frame(path_frame)
        btn_frame.pack(fill="x", padx=4, pady=4)
        ttk.Button(btn_frame, text="增加规则行", command=self._add_path_row).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="删除规则行", command=self._del_path_row).pack(side="left", padx=4)

        # ---- 特殊规则 Stop 区 ----
        stop_frame = ttk.LabelFrame(self.root, text="特殊规则指定Stop")
        stop_frame.pack(fill="x", padx=8, pady=4)

        # All Stop 勾选
        chk_frame = ttk.Frame(stop_frame)
        chk_frame.pack(fill="x", padx=4, pady=2)
        self.all_stop_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(chk_frame, text="All Stop", variable=self.all_stop_var).pack(side="left", padx=4)

        # Stop 名称行头
        stop_header = ttk.Frame(stop_frame)
        stop_header.pack(fill="x", padx=4, pady=2)
        ttk.Label(stop_header, text="名称包含字段（连续字符串，不区分大小写）", width=40).pack(side="left", padx=2)

        # Stop 名称行容器
        self.stop_row_container = ttk.Frame(stop_frame)
        self.stop_row_container.pack(fill="x", padx=4, pady=2)

        # Stop 增删按钮
        stop_btn_frame = ttk.Frame(stop_frame)
        stop_btn_frame.pack(fill="x", padx=4, pady=4)
        ttk.Button(stop_btn_frame, text="增加规则行", command=self._add_stop_row).pack(side="left", padx=4)
        ttk.Button(stop_btn_frame, text="删除规则行", command=self._del_stop_row).pack(side="left", padx=4)

        # ---- 创建按钮 ----
        create_frame = ttk.Frame(self.root)
        create_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(create_frame, text="创建 Event", command=self._create_events).pack(side="left", padx=4)
        ttk.Button(create_frame, text="清空日志", command=self._clear_log).pack(side="left", padx=4)

        # ---- 日志区 ----
        log_frame = ttk.LabelFrame(self.root, text="日志")
        log_frame.pack(fill="both", expand=True, padx=8, pady=4)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, state="disabled",
                                                   wrap="word", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=4, pady=4)

    # ===================== 路径行管理 =====================

    def _init_rows(self):
        """初始化至少一行"""
        self._add_path_row()
        self._add_stop_row()

    def _add_path_row(self):
        """增加一行路径映射"""
        row_frame = ttk.Frame(self.row_container)
        row_frame.pack(fill="x", pady=1)

        am_entry = ttk.Entry(row_frame, width=30)
        am_entry.pack(side="left", padx=2)
        am_btn = ttk.Button(row_frame, text="获取路径", width=8,
                            command=lambda e=am_entry: self._fetch_path(e))
        am_btn.pack(side="left", padx=2)

        ev_entry = ttk.Entry(row_frame, width=30)
        ev_entry.pack(side="left", padx=2)
        ev_btn = ttk.Button(row_frame, text="获取路径", width=8,
                            command=lambda e=ev_entry: self._fetch_path(e))
        ev_btn.pack(side="left", padx=2)

        self.actor_mixer_entries.append(am_entry)
        self.event_entries.append(ev_entry)

    def _del_path_row(self):
        """删除最后一行路径映射（最少保留一行）"""
        if len(self.actor_mixer_entries) <= 1:
            self._log("路径映射最少保留一行")
            return
        row_frame = self.actor_mixer_entries[-1].master
        row_frame.destroy()
        self.actor_mixer_entries.pop()
        self.event_entries.pop()

    def _add_stop_row(self):
        """增加一行 Stop 名称规则"""
        row_frame = ttk.Frame(self.stop_row_container)
        row_frame.pack(fill="x", pady=1)

        name_entry = ttk.Entry(row_frame, width=30)
        name_entry.pack(side="left", padx=2)
        name_btn = ttk.Button(row_frame, text="获取名称", width=8,
                              command=lambda e=name_entry: self._fetch_name(e))
        name_btn.pack(side="left", padx=2)

        self.stop_name_entries.append(name_entry)

    def _del_stop_row(self):
        """删除最后一行 Stop 名称规则（最少保留一行）"""
        if len(self.stop_name_entries) <= 1:
            self._log("Stop规则最少保留一行")
            return
        row_frame = self.stop_name_entries[-1].master
        row_frame.destroy()
        self.stop_name_entries.pop()

    # ===================== WAAPI 连接与获取 =====================

    def _connect(self):
        """创建新的 WAAPI 连接"""
        client = WaapiClient("ws://127.0.0.1:8080/waapi")
        if not client.is_connected():
            raise RuntimeError("无法连接到 Wwise，请确保 Wwise 已启动且 WAAPI 已启用")
        return client

    def _fetch_path(self, entry_widget):
        """获取 Wwise 选中对象的路径填入输入框"""
        client = None
        try:
            client = self._connect()
            result = client.call(
                "ak.wwise.ui.getSelectedObjects",
                options={"return": ["path"]}
            )
            objects = result.get("objects", [])
            if objects:
                path = objects[0].get("path", "")
                entry_widget.delete(0, "end")
                entry_widget.insert(0, path)
                self._log(f"获取路径: {path}")
            else:
                self._log("未选中任何对象")
        except Exception as e:
            self._log(f"获取路径失败: {e}")
        finally:
            if client:
                client.disconnect()

    def _fetch_name(self, entry_widget):
        """获取 Wwise 选中对象的名称填入输入框"""
        client = None
        try:
            client = self._connect()
            result = client.call(
                "ak.wwise.ui.getSelectedObjects",
                options={"return": ["name"]}
            )
            objects = result.get("objects", [])
            if objects:
                name = objects[0].get("name", "")
                entry_widget.delete(0, "end")
                entry_widget.insert(0, name)
                self._log(f"获取名称: {name}")
            else:
                self._log("未选中任何对象")
        except Exception as e:
            self._log(f"获取名称失败: {e}")
        finally:
            if client:
                client.disconnect()

    # ===================== 核心：创建 Event =====================

    def _create_events(self):
        """执行 Event 创建逻辑"""
        # 验证输入
        pairs = []
        for am_entry, ev_entry in zip(self.actor_mixer_entries, self.event_entries):
            am_path = am_entry.get().strip()
            ev_path = ev_entry.get().strip()
            if not am_path or not ev_path:
                self._log("ActorMixer路径和Event路径不能为空")
                return
            pairs.append((am_path, ev_path))

        # 获取 Stop 名称关键词列表
        stop_keywords = []
        for entry in self.stop_name_entries:
            kw = entry.get().strip()
            if kw:
                stop_keywords.append(kw.lower())

        all_stop = self.all_stop_var.get()

        client = None
        try:
            client = self._connect()

            # 开始 Undo Group
            client.call("ak.wwise.core.undo.beginGroup")

            total_play = 0
            total_stop = 0
            skipped = 0

            for am_path, ev_path in pairs:
                # 获取 ActorMixer 下第一层子对象
                children = self._get_direct_children(client, am_path)
                if not children:
                    self._log(f"[{am_path}] 无子对象")
                    continue

                for child in children:
                    child_name = child["name"]
                    child_path = child["path"]

                    # 检查是否已被引用
                    if self._is_referenced(client, child["id"]):
                        self._log(f"跳过已引用: {child_name}")
                        skipped += 1
                        continue

                    # 判断是否需要创建 Stop
                    need_stop = False
                    if all_stop:
                        need_stop = True
                    elif stop_keywords:
                        name_lower = child_name.lower()
                        for kw in stop_keywords:
                            if kw in name_lower:
                                need_stop = True
                                break

                    # 创建 Play Event
                    play_event_name = f"Play_{child_name}"
                    play_args = {
                        "parent": ev_path,
                        "name": play_event_name,
                        "type": "Event",
                        "onNameConflict": "merge",
                        "children": [
                            {"name": "", "type": "Action", "@ActionType": 1, "@Target": child_path}
                        ]
                    }
                    play_result = client.call("ak.wwise.core.object.create", play_args)
                    if play_result:
                        total_play += 1
                        self._log(f"✅ Play: {play_event_name}")
                    else:
                        self._log(f"❌ Play创建失败: {play_event_name}")

                    # 创建 Stop Event（独立 Event，不在同一 children 中）
                    if need_stop:
                        stop_event_name = f"Stop_{child_name}"
                        stop_args = {
                            "parent": ev_path,
                            "name": stop_event_name,
                            "type": "Event",
                            "onNameConflict": "merge",
                            "children": [
                                {"name": "", "type": "Action", "@ActionType": 2,
                                 "@Target": child_path, "@Scope": 1, "@FadeTime": 0.8}
                            ]
                        }
                        stop_result = client.call("ak.wwise.core.object.create", stop_args)
                        if stop_result:
                            total_stop += 1
                            self._log(f"✅ Stop: {stop_event_name} (Global, FadeTime:0.8)")
                        else:
                            self._log(f"❌ Stop创建失败: {stop_event_name}")

            # 结束 Undo Group
            client.call("ak.wwise.core.undo.endGroup", {"displayName": "批量创建Event"})
            self._log(f"\n========== 完成 ==========")
            self._log(f"Play Event: {total_play} | Stop Event: {total_stop} | 跳过已引用: {skipped}")

        except Exception as e:
            self._log(f"创建Event失败: {e}")
        finally:
            if client:
                client.disconnect()



    def _get_direct_children(self, client, parent_path):
        """获取父路径下第一层子对象（name, id, type, path）"""
        result = []
        try:
            res = client.call(
                "ak.wwise.core.object.get",
                {"waql": f'$ "{parent_path}"'},
                options={"return": ["children.id", "children.type", "children.name", "children.path"]}
            )
            if not res or not res.get("return"):
                return result
            for item in res["return"]:
                child_ids = item.get("children.id", [])
                child_types = item.get("children.type", [])
                child_names = item.get("children.name", [])
                child_paths = item.get("children.path", [])
                for cid, ctype, cname, cpath in zip(child_ids, child_types, child_names, child_paths):
                    result.append({"id": cid, "name": cname, "type": ctype, "path": cpath})
        except Exception as e:
            self._log(f"获取子对象失败: {e}")
        return result

    def _is_referenced(self, client, object_id):
        """检查对象是否已被引用（通过 WAQL referencesTo 查询）"""
        try:
            waql = f'$ "{object_id}" select referencesTo'
            res = client.call(
                "ak.wwise.core.object.get",
                {"waql": waql},
                options={"return": ["id", "type"]}
            )
            if not res or not res.get("return"):
                return False
            return len(res["return"]) > 0
        except Exception:
            return False

    # ===================== 日志 =====================

    def _log(self, msg):
        """向日志区输出消息"""
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _clear_log(self):
        """清空日志区"""
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    # ===================== 运行 =====================

    def run(self):
        """启动 UI 应用"""
        self.root.mainloop()


if __name__ == "__main__":
    app = EventCreatorApp()
    app.run()