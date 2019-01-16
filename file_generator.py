import os


class File:
    def __init__(self, filename):
        if os.path.isfile(f'./macros/{filename}'):
            raise Exception('File already exists')

        self.file = open(f'./macros/{filename}', "w")

        self.file.write("\n")
        self.file.write("Sub AutoTextCongress()")
        self.file.write("\n\n")
        self.file.write("Dim oAutoText As AutoTextEntry")
        self.file.write("\n\n")

    def add_autotext(self, snippet, value):
        self.file.write('Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries.Add(Name:="{}", Range:=Selection.Range)'.format(snippet))
        self.file.write('\n\t')
        self.file.write('oAutoText.Value = "{}"'.format(value))
        self.file.write("\n")

    def close_file(self):
        self.file.write("Set oAutoText = Nothing")
        self.file.write('\n\n')
        self.file.write("End Sub\n")
        self.file.close()

