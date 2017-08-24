import os, sys
import wx
from layouts import layout_main
from controllers import DataFetcher
import _thread

if os.path.isdir("output") == False: 
	os.mkdir("output")

class AppFrame(layout_main.MainFrame):
	def __init__(self, parent):
		layout_main.MainFrame.__init__(self, parent)
	
	def onClickStart(self, event):
		print("Preparing Spreadsheet...")
		inputFilename = self.inputFilePicker.GetPath()
		outputFilename = self.outputFilename.GetValue()
		print(inputFilename)
		print(outputFilename)
		fetcher = DataFetcher.DataFetcher()
		try:
			_thread.start_new_thread(fetcher.fetchData, (inputFilename, outputFilename))
		except:
			print("Unable to start a thread")
			
app = wx.App(False)
frame = AppFrame(None)
frame.Show(True)
app.MainLoop()
