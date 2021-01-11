# Copyright 2021 Dawid Pieper
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import addonHandler
import globalPluginHandler
import scriptHandler
import ui
import api
import sys
import os
import ctypes
import ctypes.wintypes
import time

addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = "Pinfo"
	class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
		_fields_ = [
			('cb', ctypes.wintypes.DWORD),
			('PageFaultCount', ctypes.wintypes.DWORD),
			('PeakWorkingSetSize', ctypes.c_size_t),
			('WorkingSetSize', ctypes.c_size_t),
			('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
			('QuotaPagedPoolUsage', ctypes.c_size_t),
			('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
			('QuotaNonPagedPoolUsage', ctypes.c_size_t),
			('PagefileUsage', ctypes.c_size_t),
			('PeakPagefileUsage', ctypes.c_size_t),
			('PrivateUsage', ctypes.c_size_t),
		]

	Psapi = ctypes.WinDLL('Psapi.dll')
	GetProcessMemoryInfo = Psapi.GetProcessMemoryInfo
	GetProcessMemoryInfo.restype = ctypes.wintypes.BOOL
	GetProcessMemoryInfo.argtypes = [ctypes.wintypes.HANDLE, ctypes.POINTER(PROCESS_MEMORY_COUNTERS_EX), ctypes.wintypes.DWORD]
	Kernel32 = ctypes.WinDLL('kernel32.dll')
	OpenProcess = Kernel32.OpenProcess
	OpenProcess.restype = ctypes.wintypes.HANDLE
	QueryFullProcessImageName = Kernel32.QueryFullProcessImageNameW
	QueryFullProcessImageName.restype = ctypes.wintypes.DWORD
	GetProcessTimes = Kernel32.GetProcessTimes
	GetProcessTimes.restype = ctypes.wintypes.DWORD
	CloseHandle = Kernel32.CloseHandle
	MAX_PATH = 260
	PROCESS_QUERY_INFORMATION = 0x0400
	UNIXTIMEDIF = -11644473600

	def formatSize(self,sz):
		if sz<1024: return str(sz)+"B"
		elif sz<1024**2: return str(round(sz/1024, 1))+"kB"
		elif sz<1024**3: return str(round(sz/1024**2, 1))+"MB"
		else: return str(round(sz/1024**4, 1))+"GB"

	@scriptHandler.script(
		description=_("Press once to read current process path, RAM and CPU utilization, press twice to copy process path to clipboard."),
		gesture="KB:NVDA+shift+f1"
	)
	def script_pinfo(self,gesture):
		focus=api.getFocusObject()
		if(focus is None): return
		pid=focus.processID
		hProcess = self.OpenProcess(self.PROCESS_QUERY_INFORMATION, False, pid)
		if hProcess==0: return
		readInfo=""
		path = (ctypes.c_wchar*self.MAX_PATH)()
		pathSize=ctypes.wintypes.DWORD(260)
		if self.QueryFullProcessImageName(hProcess, 0, path, ctypes.byref(pathSize))>0:
			readInfo+=path.value
		if scriptHandler.getLastScriptRepeatCount() > 0:
			if api.copyToClip(readInfo): readInfo = _("Process path copied to clipboard.")
		else:
			creationTime = ctypes.c_ulonglong()
			exitTime = ctypes.c_ulonglong()
			kernelTime = ctypes.c_ulonglong()
			userTime = ctypes.c_ulonglong()
			if self.GetProcessTimes(hProcess, ctypes.byref(creationTime), ctypes.byref(exitTime), ctypes.byref(kernelTime), ctypes.byref(userTime))!=0:
				runningTime = time.time()-(creationTime.value/10000000+self.UNIXTIMEDIF)
				cpuUsage = (((kernelTime.value+userTime.value)/10000000)/runningTime)/os.cpu_count()*100
				readInfo+="\n"+_("Cpu usage")+": "+str(round(cpuUsage, 1))+"%"
			mem = self.PROCESS_MEMORY_COUNTERS_EX()
			if self.GetProcessMemoryInfo(hProcess, ctypes.byref(mem), ctypes.sizeof(mem)):
				readInfo+="\n"+_("Memory usage")+": "+self.formatSize(mem.WorkingSetSize)
				readInfo+="\n"+_("Peak memory usage")+": "+self.formatSize(mem.PeakWorkingSetSize)
			self.CloseHandle(hProcess)
		ui.message(readInfo)