# DefenderBug

Many end-users raise drops of icons from Start Menu, Taskbar and Desktop.

The problem is behind a wrong Windows Defender signature (1.381.2140). This signature tells to Defender to remove all icons because not secure, even the application is Edge, Word, Excel or whatever.

The rule under investigation is "Block Win32 API calls from Office macros". You must to set the rule 92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b from Enable to Audit and update the signature to 1.381.2152.

But in the meantime the links are gone! This Repro is to create a public content to restore the most known programs to Start Menu.
