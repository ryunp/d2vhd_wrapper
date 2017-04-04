# d2vhd_wrapper
## Automation wrapper for SysInternals Disk2Vhd

Disk2vhd lacks non-drive letter arguments in command line usage. This AutoHotkey program aims to fill that gap.

The biggest difference between this and Disk2Vhd's command line interface is volume selection behavior. This program not only allows drive letters, but volume labels as search terms.

This means you can now backup arbitrary volumes, such as C:\ and "System Reserved", since the system drive won't have a drive letter.

### Usage
[Disk2Vhd.exe](https://technet.microsoft.com/en-us/sysinternals/ee656415.aspx) needs to be in the same directory as this wrapper program.

    d2vhd_wrapper.exe [SWITCHES] [TERMS] OUTPUT_FILE

The logic is broken into multiple steps:
- check for preconditions
- gather volume information
- apply given parameters against volume data
- backup via Disk2Vhd gui
- teardown

### Switches
Switches are a character, or combination of characters, preceded with a forward slash.

Switch|Description
---|---
/x|Do not use Vhdx
/s|Do not use Volume Shadow Copy
/t|Test mode (skip backup step, no output file required)
/d|Debug mode (show debug info panel)
/p|Keep GUI open after backup completion
/?|Usage text

Switches control GUI options. Flagging a switch for vhdx or shadow copy mode will disable that functionality.

Switches can be used discretely or in combination, ie: '/s /d' or '/sd'.

Test mode will skip the backup step. This allows confirmation of volume selections before actual backups are run.

Debug mode will create a gui panel with internal state, term matching results, and volume info.

### Output File
A token ending with '.vhd[x]' is taken as the output file.

### Volume Selection Terms
Volume selection terms are anything not parsed as a switch (starting with '/') or output file (ending with '.vhd[x]'). If no terms are given, default selection terms are used.

Term|Description
---|---
\<none\>|Defaults to c:\ and "System Reserved"
\*|Select all volumes
\<any string\>|RegEx'd against VOLUME and LABEL fields

Zero, one, or multiple terms can be specified, not specifying any terms will fall back to defaults. An asterisk will select all available volumes.

Terms are matched against volume Name and Label columns in Disk2Vhd's volume listview.

### Example Program Usage

Open debug info panel and perform a test mode default term selection:

    d2vhd_wrapper.exe /dt

Test run to make sure selection(s) work as expected:

    d2vhd_wrapper.exe /t "OS Disk"

Just 'System Reserved' without closing the GUI after completion:

    d2vhd_wrapper.exe /p "System Reserved" d:\test.vhd

All volumes:

    d2vhd_wrapper.exe * d:\allVolumeBackup.vhdx

Default volumes:

    d2vhd_wrapper.exe e:\no\spaces\default_bup.vhdx

Default volumes (must be explicit) plus another 'Videos' volume:

    d2vhd_wrapper.exe Videos c:\ "System Reserved" e:\multipleVolumes.vhdx

Volume 'Z:' and 'adultlabel' volumes without using vhdx and shadow volume copy:

    d2vhd_wrapper.exe /x /s z:\ adultlabel "f:\spaces\need quotes\specificVolumes.vhd"