# jsonxlsx

Transforms a standard JSON format to a Excel xlsx document

## Install

A prerequisite is a the xlsx fork at https://github.com/danchoi/xlsx, which 
must be added to the cabal sandbox:

    cabal sandbox add-source ../path-to-xlsx-fork

Then you can cabal build this project in a cabal sandbox.

## Usage

```
Usage: jsonxlsx [-a DELIM] FIELDS OUTFILE [--debug]
  Transform JSON object steam to XLSX. On STDIN provide an input stream of
  newline-separated JSON objects.
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

This is the command:

```
< sample.json jsonxlsx \
  'title:"Movie Title" year stars.name:"Movie Actors" ratings.imdb:"IMDB Score"' \
  movies.xlsx 
```

This is the output (movies.xlsx):

![screen](https://raw.githubusercontent.com/danchoi/jsonxlsx/master/jsonxlsxscreen.png)


## TODO

The [xlsx](https://github.com/danchoi/xlsx/blob/master/src/Codec/Xlsx/Types.hs)
that  `jsonxlsx` uses to generate Excel output seems to have ways to control
the column widths and other characteristics of the Excel spreadsheet.
Contributors are welcome to adding ways to manipulating those controls via `jsonxlsx`.

