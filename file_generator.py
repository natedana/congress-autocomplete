import os


class BasFile:
    def __init__(self, filename, force=False):
        if os.path.isfile(f'./macros/{filename}') and not force:
            raise Exception('File already exists')

        self.file = open(f'./macros/{filename}', "w")
        self.fileString = ''

        self.end_proc(False)
        self.start_proc("Proc0")

        self.proc_count = 0

    def line(self, text=None, extra_lines=1):
        if text and isinstance(text, str):
            self.fileString += text
        self.fileString += "".join(["\n" for i in range(0, extra_lines)])

    def start_proc(self, name, vars=True):
        self.line("Sub {}()".format(name), 2)
        if vars:
            self.line("Dim oAutoText As AutoTextEntry", 2)

        self.line_count = 0

    def end_proc(self, vars=True):
        if vars:
            self.line()
            self.line("Set oAutoText = Nothing", 2)
        self.line("End Sub", 2)

    def reset_proc(self):
        self.end_proc()

        self.proc_count += 1
        proc_name = 'Proc{}'.format(self.proc_count)
        self.start_proc(proc_name)

    def add_autotext(self, snippet, value):
        if self.line_count >= 100:
            self.reset_proc()
        else:
            self.line_count += 1

        self.line('Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries.Add(Name:="{}", Range:=Selection.Range)'.format(snippet))
        self.line('\toAutoText.Value = "{}"'.format(value))

    def close_file(self):
        self.end_proc()
        self.file.write("Sub AutoTextCongress()\n\n")

        for i in range(0, self.proc_count+1):
            self.file.write(f'Call Proc{i}\n')
        self.file.write("\n"+self.fileString)

        self.file.seek(0)
        self.file.close()

