# Xat

Prints xlsx file on the standard output like `cat` command.

## Usage

```
xat [-s|--separator=<output separator>] [--print-row-num] [--print-empty-row] <file> [<sheet>]
```

![Sample excel data](https://user-images.githubusercontent.com/30518877/56023471-dc814a80-5d48-11e9-8a50-d05809c976fa.png)

```sh
$ xat /path/to/file.xlsx
    Fruit   Color   Count
    Apple   Red     10
    Orange  Orange  20
    Grape   Purple  30
    Lemon   Yellow  40
```

## License

[MIT](https://github.com/niumlaque/xat/blob/master/LICENCE)
