# jsonxlsx

Transforms a standard JSON format to a Excel xlsx document.

This project is experimental; jsonxlsx's xlsx output has not been tested with
Excel, only with OpenOffice.

## Install

You must have a recente version of the Haskell Platform on your system.

A prerequisite is a the xlsx fork at https://github.com/danchoi/xlsx, which 
must be added to the cabal sandbox before building `jsonxlsx`.  Here are the steps

    git clone git@github.com/danchoi/xlsx
    git clone git@github.com/danchoi/jsonxlsx
    cd jsonxlsx
    cabal update
    cabal sandbox init
    cabal sandbox add-source ../xlsx
    cabal install --only-dependencies
    cabal build

    # Now copy dist/build/jsonxlsx/jsonxlsx to a location on your PATH

## Usage

```
jsonxlsx

Usage: jsonxlsx [-a DELIM] FIELDS OUTFILE [--debug]
  Transform JSON object steam to XLSX. On STDIN provide an input stream of
  newline-separated JSON objects.

Available options:
  -h,--help                Show this help text
  -a DELIM                 Concatentated array elem delimiter. Defaults to
                           comma.
  FIELDS                   JSON keypath expressions
  OUTFILE                  Output file to write to. Use '-' to emit binary xlsx
                           data to STDOUT.
  --debug                  Debug keypaths
```

## Example: 

This is the input JSON object stream:

sample.json:

```json
{
  "title": "Terminator 2: Judgement Day",
  "year": 1991,
  "stars": [
    {
      "name": "Arnold Schwarzenegger"
    },
    {
      "name": "Linda Hamilton"
    }
  ],
  "ratings": {
    "imdb": 8.5
  }
}
{
  "title": "Interstellar",
  "year": 2014,
  "stars": [
    {
      "name": "Matthew McConaughey"
    },
    {
      "name": "Anne Hathaway"
    }
  ],
  "ratings": {
    "imdb": 8.9
  }
}
```

This is the command to turn this stream into an XLSX file:

```
< sample.json jsonxlsx \
  'title:"Movie Title" year stars.name:"Movie Actors" ratings.imdb:"IMDB Score"' \
  movies.xlsx 
```

This is the output (movies.xlsx):

![screen](https://raw.githubusercontent.com/danchoi/jsonxlsx/master/jsonxlsxscreen.png)

See the [README for jsontsv](https://github.com/danchoi/jsontsv) to see how to use `jq` 
to generate JSON object streams from nested JSON data structures.


## Future improvements

The [xlsx](https://github.com/danchoi/xlsx/blob/master/src/Codec/Xlsx/Types.hs)
library that  `jsonxlsx` uses to generate Excel output seems to have ways to control
the column widths and other characteristics of the Excel spreadsheet.
Contributors are welcome to adding ways to manipulating those controls via `jsonxlsx`.

## Related

* [jsontsv](https://github.com/danchoi/jsontsv)
* [table](https://github.com/danchoi/table)
* [jsonsql](https://github.com/danchoi/jsonsql)
* [tsvsql](https://github.com/danchoi/tsvsql)

 
