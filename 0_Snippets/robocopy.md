# Robocopy Snippets

Use roboopy to move and copy files.

## Move directory structure

```
ROBOCOPY "source" "destination" *.pdf /S /MOV
```

Use `"` for paths with spaces and non-latin symbols. 

`*.pdf` - files to move.

