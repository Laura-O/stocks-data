# Earning sheet 

This small tool creates stock earning sheets using data from [FMP](https://site.financialmodelingprep.com/).

## Setup

You need a FMP starter account for accessing the API. Sadly, the free API does not offer all values needed in the sheet. Add the key in the `.env` file.

```
chmod +x main.py
```

## Usage

Find the symbol of your stock and run

```
./main.py load PLTR
```

The Excel sheet will be saved in the `/data` folder.