# d2vhd_wrapper
## Automation wrapper for SysInternals Disk2Vhd

Disk2vhd lacks non-drive letter arguments in command line usage. This AutoHotkey program aims to fill that gap.

The biggest difference between this and Disk2Vhd's command line interface is volume selection behavior. This program not only allows drive letters, but volume labels as search terms.

This means you can now backup arbitrary volumes, such as C:\ and "System Reserved", since the system drive won't have a drive letter.

### Usage
[Disk2Vhd.exe](https://technet.microsoft.com/en-us/sysinternals/ee656415.aspx) needs to be in the same directory as this wrapper program.

    d2vhd_wrapper.exe [SWITCHES] [TERMS] OUTPUT_FILE

### Switches
Switches are a character, or combination of characters, preceded with a forward slash.

Switch|Description
---|---
/x|Do not use Vhdx
/s|Do not use Volume Shadow Copy
/t|Test mode (skip backup step, no output file required)
/d|Debug mode (show debug info panel)
/?|Usage text

Switches control GUI options. Flagging a switch for vhdx or shadow copy mode will disable that functionality.

Switches can be used discretely or in combination, ie: '/s /d' or '/sd'.

Test mode will skip the backup step. This allows confirmation of drive selections before actual backups are run.

Debug mode will create a gui panel with internal state, term matching results, and volume info.

### Search Terms
Search terms are anything not parsed as a switch or output file.

Term|Description
---|---
\<none\>|Defaults to c:\ and "System Reserved"
\*|Select all volumes
\<string\>|RegEx'd against VOLUME and LABEL fields

Zero, one, or multiple terms can be specified, not specifying any terms will fall back to defaults. An asterisk will select all available drives.

Terms are matched against volume Name and Label columns in Disk2Vhd's volume listview.

### Output File
A token ending with '.vhd[x]' is taken as the output file.

### Example Program Usage
Open debug info panel and perform a test mode default term selection:

    d2vhd_wrapper.exe /dt

Test mode run of a volume labeled "OS Disk" to make sure selection works as expected:

    d2vhd_wrapper.exe /t "OS Disk"

Backup all volumes:

    d2vhd_wrapper.exe * d:\allVolumeBackup.vhdx

Backup default volumes:

    d2vhd_wrapper.exe e:\some\place\without\spaces\default_backup.vhdx

Backup drive Z: and any volumes labeled "adultlabel" without using vhdx and shadow volume copy:

    d2vhd_wrapper.exe /x /s z:\ adultlabel "f:\my backup\with spaces\adultVolumes.vhd"
