{-# LANGUAGE OverloadedStrings #-}
module Main where
import Codec.Xlsx
import qualified Data.Map.Strict as M
import qualified Data.ByteString.Lazy as L
import System.Time


-- Hackage: https://hackage.haskell.org/package/xlsx
{-

   let
       sheet = def & cellValueAt (1,2) ?~ CellDouble 42.0
                   & cellValueAt (3,2) ?~ CellText "foo"
       xlsx = def & atSheet "List1" ?~ sheet

-}

main = do
  ct <- getClockTime
  let cell = def { _cellValue = Just (CellText "hello world") }
  let cellmap = M.fromList [((1,1), cell)]
  let ws = def { _wsCells = cellmap }
  let xlsx = def { _xlSheets = M.fromList [("test", ws)] }
  L.writeFile "example.xlsx" $ fromXlsx ct xlsx

