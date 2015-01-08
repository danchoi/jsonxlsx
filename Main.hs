{-# LANGUAGE OverloadedStrings #-}
module Main where
import Codec.Xlsx
import qualified Data.Map.Strict as M
import qualified Data.ByteString.Lazy as L
import qualified Data.ByteString.Lazy.Char8 as L8
import System.Time
import System.Environment (getArgs)
import Data.Aeson

-- Hackage: https://hackage.haskell.org/package/xlsx
{-

   let
       sheet = def & cellValueAt (1,2) ?~ CellDouble 42.0
                   & cellValueAt (3,2) ?~ CellText "foo"
       xlsx = def & atSheet "List1" ?~ sheet

-}

main = do
  [outfile] <- getArgs
  ct <- getClockTime
  let cell = def { _cellValue = Just (CellText "1") }
  let cell2 = def { _cellValue = Just (CellDouble 2) }
  let cellmap = M.fromList [((1,1), cell),((1,2), cell2)]
  let ws = def { _wsCells = cellmap }
  let xlsx = def { _xlSheets = M.fromList [("test", ws)] }
  if outfile == "-"
  then L8.putStr $ fromXlsx ct xlsx
  else L.writeFile outfile $ fromXlsx ct xlsx

