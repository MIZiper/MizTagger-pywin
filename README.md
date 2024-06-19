# MizTagger-pywin

A files tag system based on pywin32.

## How to use

Install the tool via pip.

```bash
pip install miz-tagger
```

There are two ways to manage the files and tags:
via Windows file explorer context menu, or via a GUI program called "MizTagger-viewer".

**Windows explorer integration**

To integrate the service into Windows file explorer,
type `MizTagger-register.exe` after pip installation,
it will prompt for privilege escalation to modify registry.

Uninstalling the service can be done by `MizTagger-unregister.exe`.

After registration, right-click on space area of a folder in file explorer, the entries will appear.
You can manage the classes and tags, which will be used for further filtering.

*Note*: for Windows 11, the classic context menu is shadow by new version.
To show directly the classic context menu, hold `Shift` when right-clicking.

![MizTagger entries for folder](./doc/images/miztagger-folder-tag-management.png)

After tag management, a file "MizTagger.json" will appear in the folder, maintaining tag and file info.

Next step is to assign tag info to individual files.

![MizTagger set file](./doc/images/miztagger-set-file-info.png)

To filter out files with tags, use the "Quick Filter" entry of the folder.
After filtering, a temp folder will be created, and copy (hard linking) related files.

Inside the result folder, further filtering can be performed. Or use a complex filtering.

![MizTagger filter result](./doc/images/miztagger-filter-result.png)

Use "Clean Result" to exit filtering.

**MizTagger viewer**

Another way to use the tool is via a GUI dialog `MizTagger-viewer.exe`,
which provides the same functionalities.

![MizTagger viewer GUI](./doc/images/miztagger-viewer.png)

1. Browse to the folder
2. Assign tags to files
3. Filter or quick filter for the results

## Known issue

- Seems virtual environment doesn't work
- Don't know how to register the service without priviledge escalation (e.g. enable the service in user level)

<!-- ## What are the tags -->

<!-- ## Content in Windows registery -->