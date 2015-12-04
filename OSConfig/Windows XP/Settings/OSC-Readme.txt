This folder is reserved for Settings

Supported Extensions:
	BAT
	CMD
	EXE (no switches are processed)
	REG
	PS1
	VBS
	All other extensions are ignored

Supporting Naming Rules (Processed In Order):
	Files with "Sample" in the name		Skipped
	Files with "Undo" in the name		Skipped
	Files with "TSYes" in the name		Applied if C:\_SMSTaskSequence\OSConfig\OSConfig.cmd does exists
	Files with "TSNo" in the name		Applied if C:\_SMSTaskSequence\OSConfig\OSConfig.cmd does NOT exist
	Files with "Modern" in the name		Applied if the Operating System is not Windows XP
	Files with "x86" in the name		Applied if the Operating System is x86
	Files with "x64" in the name		Applied if the Operating System is x64
	Files with "RunOnce" in the name	Only applied once, then moved to C:\Windows\OSConfig\Settings\RunOnce