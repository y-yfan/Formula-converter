# -*- coding: utf-8 -*-
"""
公式转换工具 - GUI 版本
将 Word 文档中的纯文本公式替换为 OMML 格式，保留原样式。
"""
import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import filedialog, ttk


def convert_doc_to_docx(doc_path, log_func):
    """将 .doc 文件转为 .docx（使用 Word 或 WPS COM）"""
    import subprocess
    import tempfile
    import shutil
    import winreg

    # 检测可用的 Office COM
    com_name = None
    for name in ['Word.Application', 'Kwps.Application']:
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name)
            winreg.CloseKey(key)
            com_name = name
            break
        except FileNotFoundError:
            continue

    if com_name is None:
        log_func('错误：本机未安装 Microsoft Word 或 WPS Office，无法自动转换 .doc 文件')
        log_func('请手动将 .doc 另存为 .docx 格式，然后选择 .docx 文件作为输入')
        return None

    app_name = 'Word' if 'Word' in com_name else 'WPS'

    tmp_doc = os.path.join(tempfile.gettempdir(), 'input_temp.doc')
    tmp_docx = os.path.join(tempfile.gettempdir(), 'input_temp.docx')
    if os.path.exists(tmp_docx):
        os.remove(tmp_docx)
    shutil.copy2(doc_path, tmp_doc)

    log_func(f'正在通过 {app_name} COM 将 .doc 转换为 .docx...')

    doc_ps_path = tmp_doc.replace(os.sep, '/')
    docx_ps_path = tmp_docx.replace(os.sep, '/')

    ps_file = os.path.join(tempfile.gettempdir(), 'convert.ps1')
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write('try {\n')
        f.write('  $app = New-Object -ComObject ' + com_name + '\n')
        f.write('  $app.Visible = $false\n')
        f.write('  $doc = $app.Documents.Open("' + doc_ps_path + '")\n')
        f.write('  $doc.SaveAs2("' + docx_ps_path + '", 16)\n')
        f.write('  $doc.Close($false)\n')
        f.write('  $app.Quit()\n')
        f.write('  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null\n')
        f.write('  Write-Output "SUCCESS"\n')
        f.write('} catch {\n')
        f.write('  Write-Output "ERROR: $_"\n')
        f.write('}\n')

    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0  # SW_HIDE
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
            capture_output=True, timeout=120, startupinfo=si
        )
        stdout = result.stdout.decode('utf-8', errors='replace').strip()
        if 'ERROR' in stdout:
            log_func(f'{app_name} COM 转换出错：{stdout}')
            return None
    except subprocess.TimeoutExpired:
        log_func('错误：COM 转换超时')
        return None

    if not os.path.exists(tmp_docx):
        log_func(f'错误：{app_name} COM 转换失败，请手动将 .doc 另存为 .docx')
        return None

    log_func(f'.doc -> .docx 转换完成（使用 {app_name}）')
    return tmp_docx


def do_convert(input_path, output_path, log_func):
    """执行公式转换（在子线程中调用）"""
    try:
        # .doc 转 .docx
        actual_input = input_path
        if input_path.lower().endswith('.doc') and not input_path.lower().endswith('.docx'):
            actual_input = convert_doc_to_docx(input_path, log_func)
            if actual_input is None:
                return False

        # 导入核心逻辑
        log_func('正在加载转换模块...')
        if getattr(sys, 'frozen', False):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        if script_dir not in sys.path:
            sys.path.insert(0, script_dir)
        from replace_formulas import process_document

        # 直接调用，log_func 实时输出进度
        log_func('正在替换公式...')
        process_document(actual_input, output_path, log_func=log_func)

        if os.path.exists(output_path):
            log_func(f'转换完成！输出文件：{output_path}')
            return True
        else:
            log_func('转换失败，输出文件未生成')
            return False

    except Exception as ex:
        msg = str(ex)
        if 'Permission denied' in msg or '另一个程序正在使用' in msg or 'WinError 32' in msg:
            if input_path and output_path and input_path.lower() == output_path.lower():
                log_func('错误：输出文件与输入文件相同，且被其他程序占用，请关闭 Word/WPS 后重试，或选择不同的输出路径')
            elif input_path and 'WinError 32' in msg:
                log_func(f'错误：输入文件被占用，请关闭 Word/WPS 中打开的 {input_path} 后重试')
            else:
                log_func(f'错误：输出文件 {output_path} 被占用，请关闭 Word/WPS 中打开的该文件后重试')
        elif 'No such file' in msg or '找不到' in msg or 'FileNotFoundError' in msg:
            if input_path and not os.path.exists(input_path):
                log_func(f'错误：输入文件不存在：{input_path}')
            else:
                log_func(f'错误：文件路径有误，请检查输入和输出路径是否正确')
        elif 'COM' in msg or 'COMObject' in msg:
            log_func('错误：无法调用 Word/WPS，请确认已安装并重启后再试')
        elif 'timeout' in msg.lower() or 'TimeoutExpired' in msg:
            log_func('错误：操作超时，文件可能过大或 Word/WPS 响应缓慢，请重试')
        elif 'codec' in msg.lower() or 'encode' in msg.lower():
            log_func('错误：文件编码问题，请将文档另存为 .docx 格式后再试')
        else:
            log_func(f'错误：{ex}')
            import traceback
            log_func(traceback.format_exc())
        return False


class FormulaConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title('公式转换工具')
        self.root.resizable(True, True)
        self.root.minsize(520, 380)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.running = False
        self.log_queue = queue.Queue()

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(1, weight=1)

        # 输入文件
        ttk.Label(main, text='输入文件：').grid(row=0, column=0, sticky=tk.W, pady=(0, 3))
        input_frame = ttk.Frame(main)
        input_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)

        ttk.Entry(input_frame, textvariable=self.input_var).grid(row=0, column=0, sticky=tk.EW)
        ttk.Button(input_frame, text='浏览...', command=self._browse_input).grid(row=0, column=1, padx=(5, 0))

        # 输出文件
        ttk.Label(main, text='输出文件：').grid(row=2, column=0, sticky=tk.W, pady=(0, 3))
        output_frame = ttk.Frame(main)
        output_frame.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_var).grid(row=0, column=0, sticky=tk.EW)
        ttk.Button(output_frame, text='浏览...', command=self._browse_output).grid(row=0, column=1, padx=(5, 0))

        # 转换按钮
        self.convert_btn = ttk.Button(main, text='开始转换', command=self._start_convert)
        self.convert_btn.grid(row=4, column=0, columnspan=2, pady=(0, 15))

        # 日志区域
        ttk.Label(main, text='处理日志：').grid(row=5, column=0, sticky=tk.W, pady=(0, 3))
        log_frame = ttk.Frame(main)
        log_frame.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW)
        main.rowconfigure(6, weight=1)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                font=('Consolas', 9))
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title='选择输入文件',
            filetypes=[('Word 文档', '*.doc *.docx'), ('所有文件', '*.*')]
        )
        if path:
            self.input_var.set(path)
            base = os.path.splitext(path)[0]
            self.output_var.set(f'{base}_公式版.docx')

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title='选择输出文件',
            defaultextension='.docx',
            filetypes=[('Word 文档', '*.docx')]
        )
        if path:
            self.output_var.set(path)

    def _log(self, msg):
        self.log_queue.put(msg)

    def _poll_log(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break

            if msg == '__DONE__':
                self.running = False
                self.convert_btn.configure(state=tk.NORMAL)
                continue

            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + '\n')
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        self.root.after(100, self._poll_log)

    def _start_convert(self):
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()

        if not input_path:
            self._log('请选择输入文件')
            return
        if not output_path:
            self._log('请选择输出文件')
            return
        if not os.path.exists(input_path):
            self._log(f'输入文件不存在：{input_path}')
            return
        if self.running:
            return

        self.running = True
        self.convert_btn.configure(state=tk.DISABLED)
        self._log(f'输入：{input_path}')
        self._log(f'输出：{output_path}')

        def run():
            do_convert(input_path, output_path, self._log)
            self.log_queue.put('__DONE__')

        threading.Thread(target=run, daemon=True).start()


def main():
    root = tk.Tk()
    FormulaConverterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
