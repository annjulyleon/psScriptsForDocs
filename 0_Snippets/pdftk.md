# PDFTK Snippets

[Pdftk](https://www.pdflabs.com/tools/pdftk-server/) Server is useful tool for pdf manipulation with simple command line. You can see full command reference [here](https://www.pdflabs.com/docs/pdftk-man-page/) and [some examples](https://www.pdflabs.com/docs/pdftk-cli-examples/).

## Merge all pdf in current folder

```
pdftk *.pdf cat output .pdf
```

## Save single page from source.pdf and save to output.pdf

```
pdftk source.pdf cat 2 output output.pdf
```

where `2` - is page number.

