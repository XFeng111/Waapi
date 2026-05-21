from PyQt6 import QtCore, QtGui, QtWidgets
import sys
import socket
from io import StringIO
import contextlib
from waapi import WaapiClient

class Ui_Wwise_SetNotes(object):
    def setupUi(self, Wwise_SetNotes):
        Wwise_SetNotes.setObjectName("Wwise_SetNotes")
        Wwise_SetNotes.resize(831, 300)
        self.lineEdit = QtWidgets.QLineEdit(parent=Wwise_SetNotes)
        self.lineEdit.setGeometry(QtCore.QRect(20, 20, 601, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(parent=Wwise_SetNotes)
        self.pushButton.setGeometry(QtCore.QRect(740, 20, 71, 31))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(parent=Wwise_SetNotes)
        self.pushButton_2.setGeometry(QtCore.QRect(640, 20, 71, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.listView = QtWidgets.QListView(parent=Wwise_SetNotes)
        self.listView.setGeometry(QtCore.QRect(20, 71, 791, 201))
        self.listView.setObjectName("listView")
        # 设置列表视图不可编辑
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)

        self.retranslateUi(Wwise_SetNotes)
        QtCore.QMetaObject.connectSlotsByName(Wwise_SetNotes)

    def retranslateUi(self, Wwise_SetNotes):
        _translate = QtCore.QCoreApplication.translate
        Wwise_SetNotes.setWindowTitle(_translate("Wwise_SetNotes", "Wwise_SetNotes"))
        self.pushButton.setText(_translate("Wwise_SetNotes", "批量Notes"))
        self.pushButton_2.setText(_translate("Wwise_SetNotes", "清空"))

def collect_output(func):
        """装饰器函数：捕获被装饰函数中的所有print输出，返回输出内容列表"""
        def wrapper(*args, **kwargs):
            output_buffer = StringIO()  # 创建字符串缓冲区，用于捕获输出
            with contextlib.redirect_stdout(output_buffer), contextlib.redirect_stderr(output_buffer):
                # 执行被装饰的函数，所有print和错误输出都会被捕获
                func(*args, **kwargs)
            
            # 获取缓冲区内容并按行分割，过滤空行
            output_content = output_buffer.getvalue().strip().split('\n')
            return [line.strip() for line in output_content if line.strip()]
        
        return wrapper

class SetNotes():
    def __init__(self):
        pass
    
    @collect_output
    def batch_add_custom_notes(self, notes_content):
        try:
            """让用户输入备注内容，为选中对象批量添加（支持外部传入内容）"""
            # # 优先使用外部传入的内容，否则让用户输入
            # if notes_content is None:
            #     print("请输入需要批量添加的备注内容（输入完成后按回车）：")
            #     notes_content = input("> ").strip()  # 获取用户输入并去除首尾空格
            #     print(f"添加备注:{notes_content}")

            # if not notes_content:
            #     print("⚠️ 备注内容不能为空，请重新运行脚本并输入内容")
            #     return
            
            """为选中对象批量添加备注（仅处理传入的内容，不进行交互式输入）"""
            print(f"准备添加备注: {notes_content}")

            # 连接 Wwise 的 WAAPI 服务（默认端口 8080）
            with WaapiClient() as client:
                print("✅ 成功连接到 Wwise WAAPI 服务")

                opt = {
                    "return": ["id", "name", "type"]  # 只返回对象的 ID、名称、类型
                    }
                # 1. 获取当前选中的对象（通过 WAAPI 调用 ak.wwise.ui.getSelectedObjects）
                result = client.call("ak.wwise.ui.getSelectedObjects",options=opt)

                # 提取选中的对象列表
                selected_objects = result.get("objects", [])
                if not selected_objects:
                    print("⚠️ 未选中任何对象，请在 Wwise 中先选择需要添加备注的对象")
                    return

                print(f"\n📌 共选中 {len(selected_objects)} 个对象，开始批量添加 Notes...")

                # 2. 遍历选中对象，使用 setNotes 接口设置备注
                success_count = 0
                fail_count = 0
                fail_details = []

                for obj in selected_objects:
                    obj_id = obj["id"]
                    obj_name = obj["name"]
                    obj_type = obj["type"]

                    try:
                        args = {
                            "object": obj_id,  # 对象的 GUID 或路径
                            "value": notes_content  # 要设置的备注内容
                            }
                        # 调用 ak.wwise.core.object.setNotes 接口
                        client.call("ak.wwise.core.object.setNotes", args)
                        success_count += 1
                        print(f"✅ 成功：[{obj_type}] {obj_name} 备注：{notes_content}")
                    except Exception as e:
                        fail_count += 1
                        fail_details.append(f"❌ 失败：[{obj_type}] {obj_name}（错误：{str(e)}）")

                # 输出统计结果
                print(f"\n📊 操作完成：成功 {success_count} 个，失败 {fail_count} 个")
                if fail_details:
                    print("\n❌ 失败详情：")
                    for detail in fail_details:
                        print(detail)

        except (ConnectionRefusedError, socket.error) as e:
            print("❌ 无法连接到 Wwise WAAPI 服务，请确保：")
            print("1. Wwise 已启动")
            print("2. WAAPI 服务已开启（在 Wwise 设置中确认）")
            print("3. 端口未被占用（默认端口 8080）")
        except Exception as e:
            print(f"❌ 发生错误：{str(e)}")

    # if __name__ == "__main__":
    #     print("===== Wwise 批量添加备注工具 =====")
    #     # 调用函数并获取所有输出内容
    #     output_lines = batch_add_custom_notes()
        
    #     # 演示：打印收集到的输出（实际使用时可根据需求处理）
    #     print("\n===== 收集到的输出内容 =====")
    #     for line in output_lines:
    #         print(line)
        
    #     input("\n按回车键退出...")

class MainWindow(QtWidgets.QMainWindow, Ui_Wwise_SetNotes):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        # 初始化列表模型
        self.model = QtGui.QStandardItemModel()
        self.listView.setModel(self.model)
        # 绑定按钮事件
        self.pushButton.clicked.connect(self.batch_add_custom_notes)
        self.pushButton_2.clicked.connect(self.reset_all)

    def Note(self):
        """获取lineEdit中的输入内容"""
        return self.lineEdit.text().strip()  # 返回输入框内容（去除首尾空格）

    def reset_all(self):
        """重置输入框和列表视图"""
        self.lineEdit.clear()
        self.model.clear()

    def add_log(self, text, is_error=False):
        """向列表视图添加日志信息"""
        item = QtGui.QStandardItem(text)
        # 错误信息显示为红色
        if is_error:
            item.setForeground(QtGui.QColor(255, 0, 0))
        self.model.appendRow(item)
        # 自动滚动到底部
        self.listView.scrollToBottom()

    def batch_add_custom_notes(self):
        """调用SetNotes中的批量添加函数，并显示结果"""
        # 清空之前的日志
        self.model.clear()
        # 获取输入内容
        notes_content = self.Note()
        # 调用SetNotes中的函数并获取输出
        try:
            set_notes = SetNotes()
            # 调用带装饰器的函数，获取所有输出行
            output_lines = set_notes.batch_add_custom_notes(notes_content=notes_content)
            # 显示输出结果
            for line in output_lines:
                # 判断是否为错误信息（包含特定标记）
                is_error = "⚠️" in line or "❌" in line
                self.add_log(line, is_error)
        except Exception as e:
            self.add_log(f"调用函数失败：{str(e)}", is_error=True)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
