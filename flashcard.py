import wx
import wx.lib.agw.infobar
import sys
import datetime
import random
import os
import tempfile

try:
    # pip install gTTS pygame
    from gtts import gTTS
    from pygame import mixer
    mixer.init()
    HAS_TTS = True
except ImportError:
    HAS_TTS = False

assert sys.version_info > (3,), "We used python 3 super() syntax"

COLOURS = "#F6EFF7 #D0D1E6 #A6BDDB #67A9CF #3690C0 #02818A #016450".split()
DEFAULTWIDTH = 1000
DEFAULTHEIGHT = 500

def initwx():
    wx.DisableAsserts()
    return wx.App()

class WrapStaticText(wx.lib.agw.infobar.AutoWrapStaticText):
    """Override the constructor of AutoWrapStaticText to allow centre aligned text"""
    def __init__(self, *args, **kwargs):
        for k, v in zip(["parent", "ID", "label", "pos", "size", "style", "name"], args):
            if k not in kwargs:
                kwargs[k] = v
        kwargs["style"] = kwargs.get("style", 0) | wx.ST_NO_AUTORESIZE
        wx.lib.stattext.GenStaticText.__init__(self, **kwargs)
        self.label = self.GetLabel()
        self.SetOwnForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_INFOTEXT))
        self.Bind(wx.EVT_SIZE, self.OnSize)

class FlashCard(wx.Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.qbank = []

        # timer for stat display
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.OnTimer)

        # geometry
        self.SetTitle("Flashcard App")
        self.Centre()

        # menu bar
        menubar = wx.MenuBar()
        filemenu = wx.Menu()
        openitem = filemenu.Append(wx.ID_OPEN, "&Open", "Open file")
        self.Bind(wx.EVT_MENU, self.OnOpen, openitem)
        filemenu.AppendSeparator()
        quititem = filemenu.Append(wx.ID_EXIT, "&Quit", "Quit app")
        self.Bind(wx.EVT_MENU, self.OnQuit, quititem)
        menubar.Append(filemenu, "&File")
        self.SetMenuBar(menubar)

        # status bar
        self.statbar = self.CreateStatusBar(2)

        # widgets: master panel
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour(COLOURS[0])

        # widgets: left and right panels, children of master panel
        displayfont = wx.Font(18, wx.DECORATIVE, wx.NORMAL, wx.NORMAL, False, "Arial")
        self.lpan = wx.Panel(self.panel)
        self.rpan = wx.Panel(self.panel)
        self.ltxt = WrapStaticText(self.lpan, style=wx.ALIGN_CENTRE_HORIZONTAL)
        self.ltxt.SetFont(displayfont)
        self.rtxt = WrapStaticText(self.rpan, style=wx.ALIGN_CENTRE_HORIZONTAL)
        self.rtxt.SetFont(displayfont)
        self.lpan.SetBackgroundColour(COLOURS[1])
        self.ltxt.SetBackgroundColour(COLOURS[1])
        self.rpan.SetBackgroundColour(COLOURS[2])
        self.rtxt.SetBackgroundColour(COLOURS[2])

        # sizers to combine all widgets
        twocol = wx.GridSizer(cols=2, vgap=2, hgap=2)
        twocol.Add(self.lpan, wx.ID_ANY, wx.EXPAND|wx.ALL, 2)
        twocol.Add(self.rpan, wx.ID_ANY, wx.EXPAND|wx.ALL, 2)

        lcol = wx.BoxSizer(wx.HORIZONTAL)
        lcol.Add(self.ltxt, proportion=1, flag=wx.ALIGN_CENTER, border=15)

        rcol = wx.BoxSizer(wx.HORIZONTAL)
        rcol.Add(self.rtxt, proportion=1, flag=wx.ALIGN_CENTER, border=15)

        self.lpan.SetSizer(lcol)
        self.rpan.SetSizer(rcol)
        self.panel.SetSizer(twocol)

        # mouse and keyboard events
        self.lpan.Bind(wx.EVT_LEFT_UP, self.Next)
        self.rpan.Bind(wx.EVT_LEFT_UP, self.Next)
        self.lpan.Bind(wx.EVT_CHAR_HOOK, self.OnKeyPress)
        self.rpan.Bind(wx.EVT_CHAR_HOOK, self.OnKeyPress)

    def OnOpen(self, e):
        with wx.FileDialog(self, "Open question bank",
                           wildcard="Excel and CSV files (*.xlsx;*.csv)|*.xlsx;*.csv|"
                                    "Excel files (*.xlsx)|*.xlsx|"
                                    "CSV files|*.csv",
                           style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST
                          ) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return # nothing happened
            path = dialog.GetPath()
        if path.lower().endswith(".xlsx"):
            import openpyxl
            workbook = openpyxl.load_workbook(path, read_only=True)
            wb_sheets = [
                (name, read_questions([[cell.value for cell in row] for row in workbook[name].rows]))
                for name in workbook.sheetnames
            ]
            qbank = [(n, d) for n, d in wb_sheets if d]
            if qbank:
                self.SetTitle(qbank[0][0])
                qbank = qbank[0][1]
        elif path.lower().endswith(".csv"):
            import csv
            with open(path) as fp:
                data = list(filter(None, [row for row in csv.reader(fp)]))
            qbank = read_questions(data)
        else:
            wx.MessageDialog(None, "Unknown file extension %s" % path, "Error", wx.OK|wx.ICON_ERROR).ShowModal()
            return
        if not qbank:
            errormsg = "Cannot find questions in %s. Please make a column with header 'question' " \
                       "and a column with header 'answer' in the file for the question and answers."
            wx.MessageDialog(None, errormsg % path, "No questions", wx.OK|wx.ICON_ERROR).ShowModal()
            return
        self.qbank = qbank
        self.start = datetime.datetime.utcnow()
        self.state = 1 # 0 = Q or 1 = Q&A
        self.count = 0 # num of Q&A shown
        self.qnum = -1 # state 1: Q number to shown next; state 0: Q number shown
        self.Next()

    def OnQuit(self, e):
        self.Close()

    def OnKeyPress(self, e):
        keycode = e.GetKeyCode()
        if keycode in [wx.WXK_SPACE, wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER]:
            self.Next()
        else:
            e.Skip() # do not handle this

    def OnTimer(self, e):
        if not self.qbank:
            return
        now = datetime.datetime.utcnow()
        elapsed = (now - self.start).total_seconds()
        mins = elapsed // 60
        secs = elapsed % 60
        self.statbar.SetStatusText("%d questions done" % self.count)
        self.statbar.SetStatusText("%d:%d" % (mins, secs), 1)

    def Next(self, e=None):
        if not self.qbank:
            return
        if self.state == 0:
            # showing Q only, now reveal A
            to_say = self.qbank[self.qnum][1]
            self.rtxt.SetLabel(to_say)
            self.state = 1
            self.count += 1
        else:
            # showing Q&A, now switch to next Q
            self.qnum += 1
            if self.qnum == len(self.qbank):
                self.qnum = 0
                random.shuffle(self.qbank)
            to_say = self.qbank[self.qnum][0]
            self.ltxt.SetLabel(to_say)
            self.rtxt.SetLabel("")
            self.state = 0
        self.OnTimer(e=None) # update status bar
        self.ltxt.Wrap(self.lpan.GetSize().GetWidth())
        self.rtxt.Wrap(self.rpan.GetSize().GetWidth())
        if HAS_TTS:
            handle, filename = tempfile.mkstemp(suffix=".mp3")
            gTTS(text=to_say, lang="en").save(filename)
            mixer.music.load(filename)
            mixer.music.play()
            os.unlink(filename)

    def Show(self):
        self.timer.Start(1000) # start timer with 1 sec interval
        wx.CallLater(1, self.Next)
        super().Show()


def read_questions(array):
    header, data = array[0], array[1:]
    header = [(col or "").lower().strip() for col in header]
    if "question" in header and "answer" in header:
        qcol = next(i for i,c in enumerate(header) if c == "question")
        acol = next(i for i,c in enumerate(header) if c == "answer")
        return [[row[qcol].strip(), row[acol].strip()] for row in data]
    return []

def main():
    app = initwx()
    win = FlashCard(None)
    win.Show()
    win.SetSize((DEFAULTWIDTH, DEFAULTHEIGHT))
    app.MainLoop()

if __name__ == "__main__":
    main()
