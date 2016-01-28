{-# LANGUAGE OverloadedStrings, QuasiQuotes, BangPatterns, ScopedTypeVariables #-}
module Main where
import Codec.Xlsx
import Data.Text (Text, take)
import Control.Applicative
import qualified Data.Map.Strict as M
import qualified Data.ByteString as B
import qualified Data.ByteString.Lazy as L
import Data.ByteString.Lazy as BL hiding (map, intersperse, zip, concat)
import qualified Data.ByteString.Lazy.Char8 as L8 
import System.Time
import Data.Time.Clock.POSIX (getPOSIXTime)
import System.Environment (getArgs)
import Data.Aeson
import Data.Monoid
import Data.List (intersperse, zip)
import qualified Data.Attoparsec.Text as AT
import Data.Attoparsec.Lazy as Atto hiding (Result)
import Data.Attoparsec.ByteString.Char8 (endOfLine, sepBy)
import qualified Data.HashMap.Lazy as HM
import qualified Data.Vector as V
import Data.Scientific  (Scientific, floatingOrInteger)
import qualified Data.Text as T
import qualified Data.Text.Lazy.IO as TL
import qualified Options.Applicative as O
import Control.Monad (when)
import System.Exit
import Data.String.QQ 
import qualified Data.Text.Encoding as T (encodeUtf8, decodeUtf8)
import Test.HUnit

-- Hackage: https://hackage.haskell.org/package/xlsx
{-

   let
       sheet = def & cellValueAt (1,2) ?~ CellDouble 42.0
                   & cellValueAt (3,2) ?~ CellText "foo"
       xlsx = def & atSheet "List1" ?~ sheet

-}


data Options = Options { 
    arrayDelim :: String
  , jsonExpr :: String
  , outputFile :: String
  , debugKeyPaths :: Bool
  , maxStringLen :: Int
  } deriving Show

defaultMaxBytes :: Int
defaultMaxBytes = 32768

parseOpts :: O.Parser Options
parseOpts = Options 
  <$> O.strOption (O.metavar "DELIM" <> O.value "," <> O.short 'a' <> O.help "Concatentated array elem delimiter. Defaults to comma.")
  <*> O.argument O.str (O.metavar "FIELDS" <> O.help "JSON keypath expressions")
  <*> O.argument O.str (O.metavar "OUTFILE" <> O.help "Output file to write to. Use '-' to emit binary xlsx data to STDOUT.") 
  <*> O.switch (O.long "debug" <> O.help "Debug keypaths")
  <*> maxStrLen

opts = O.info (O.helper <*> parseOpts)
          (O.fullDesc 
            <> O.progDesc [s|Transform JSON object steam to XLSX.  
                    On STDIN provide an input stream of newline-separated JSON objects. |]
            <> O.header "jsonxlsx"
            <> O.footer "See https://github.com/danchoi/jsonxlsx for more information.")

maxStrLen :: O.Parser Int
maxStrLen = O.option O.auto
            ( O.long "maxlen"
           <> O.short 'l'
           <> O.metavar "MAXLEN"
           <> O.value defaultMaxBytes
           <> O.help "Limit the length of strings (-1 for unlimited)" )

main = do
  Options arrayDelim expr outfile debugKeyPaths maxLen <- O.execParser opts
  x <- BL.getContents 
  ct <- getPOSIXTime
  let xs :: [Value]
      xs = decodeStream x
      ks = parseKeyPath $ T.pack expr
      ks' :: [[Key]]
      ks' = [k | KeyPath k _ <- ks]
      arrayDelim' = T.pack arrayDelim
      hs :: [Text] -- header labels
      hs = map keyPathToHeader ks
  -- extract JSON
  let xs' :: [[Value]]
      xs' = map (evalToValues arrayDelim' ks') xs
      headerCells = map mkHeaderCell hs
      headerIndexedCells :: [((Int,Int), Cell)]
      headerIndexedCells = zip [(1, x) | x <- [1..]] headerCells
      rows :: [[Cell]]
      rows = map (map (jsonToCell . truncateStr maxLen)) xs'
      rowsIndexedCells :: [[((Int,Int), Cell)]]
      rowsIndexedCells = map mkRowIndexedCells $ zip [2..] rows
      allCells = concat (headerIndexedCells:rowsIndexedCells)
      cellMap :: CellMap
      cellMap = M.fromList allCells

  when debugKeyPaths $ do
     Prelude.putStrLn $ "Key Paths: " ++ show ks
     print hs
     print headerCells
     print headerIndexedCells
     exitSuccess
  let ws = def { _wsCells = cellMap }
  let xlsx = def { _xlSheets = M.fromList [("test", ws)] }
  if outfile == "-"
  then L8.putStr $ fromXlsx ct xlsx
  else L.writeFile outfile $ fromXlsx ct xlsx

mkRowIndexedCells :: (Int, [Cell]) -> [((Int, Int), Cell)]
mkRowIndexedCells (rowNumber, cells) = zip [(rowNumber, x) | x <- [1..]] cells

mkHeaderCell :: Text -> Cell
mkHeaderCell x = def { _cellValue = Just (CellText x) }

truncateStr :: Int -> Value -> Value
truncateStr (-1) v = v
truncateStr maxlen (String xs) = 
    let ellipsis = ("..." :: Text)
        maxlen' = maxlen - (bytelen ellipsis)
    in if (bytelen xs) > maxlen
       then let s = truncateText maxlen' xs 
            in String $ s <> ellipsis
       else String xs
truncateStr _ v = v

-- This should get Text approximately under the byte limit, erring on the side of being
-- too aggressive.
truncateText :: Int -> Text -> Text
truncateText maxBytes s | bytelen s > maxBytes = 
      truncateText maxBytes . T.take (T.length s - d) $ s
    where d = bytelen s - maxBytes
truncateText _ s = s

bytelen :: Text -> Int
bytelen = B.length . T.encodeUtf8 

jsonToCell :: Value -> Cell
jsonToCell (String x) = def { _cellValue = Just (CellText x) }
jsonToCell Null = def { _cellValue = Nothing }
jsonToCell (Number x) = def { _cellValue = Just (CellDouble $ scientificToDouble x) }
jsonToCell (Bool x) = def { _cellValue = Just (CellBool x) }
jsonToCell (Object _) = def { _cellValue = Just (CellText "[Object]") }
jsonToCell (Array _) = def { _cellValue = Just (CellText "[Array]") }

scientificToDouble :: Scientific -> Double
scientificToDouble x = 
    case floatingOrInteger x of
        Left float -> float
        Right int -> fromIntegral int

------------------------------------------------------------------------
-- decode JSON object stream

decodeStream :: (FromJSON a) => BL.ByteString -> [a]
decodeStream bs = case decodeWith json bs of
    (Just x, xs) | xs == mempty -> [x]
    (Just x, xs) -> x:(decodeStream xs)
    (Nothing, _) -> []

decodeWith :: (FromJSON a) => Parser Value -> BL.ByteString -> (Maybe a, BL.ByteString)
decodeWith p s =
    case Atto.parse p s of
      Atto.Done r v -> f v r
      Atto.Fail _ _ _ -> (Nothing, mempty)
  where f v' r = (\x -> case x of 
                      Success a -> (Just a, r)
                      _ -> (Nothing, r)) $ fromJSON v'

------------------------------------------------------------------------
-- JSON parsing and data extraction

-- | KeyPath may have an alias for the header output
data KeyPath = KeyPath [Key] (Maybe Text) deriving Show

data Key = Key Text | Index Int deriving (Eq, Show)

parseKeyPath :: Text -> [KeyPath]
parseKeyPath s = case AT.parseOnly pKeyPaths s of
    Left err -> error $ "Parse error " ++ err 
    Right res -> res

keyPathToHeader :: KeyPath -> Text
keyPathToHeader (KeyPath _ (Just alias)) = alias
keyPathToHeader (KeyPath ks Nothing) = 
    mconcat $ intersperse "." $ [x | Key x <- ks] -- exclude Index keys

spaces = many1 AT.space

pKeyPaths :: AT.Parser [KeyPath]
pKeyPaths = pKeyPath `AT.sepBy` spaces

pKeyPath :: AT.Parser KeyPath
pKeyPath = KeyPath 
    <$> (AT.sepBy1 pKeyOrIndex (AT.takeWhile1 $ AT.inClass ".["))
    <*> (pAlias <|> pure Nothing)

------------------------------------------------------------------------
-- | A column header alias is designated by : followed by alphanum string after keypath
-- The alias string may be quoted with double quotes if it contains strings.

pAlias :: AT.Parser (Maybe Text)
pAlias = do
    AT.char ':'
    alias <- (AT.takeWhile1 (AT.inClass "a-zA-Z0-9_-"))  <|> quotedString 
    return . Just $ alias

quotedString = AT.char '"' *> (AT.takeWhile1 (AT.notInClass "\"")) <* AT.char '"'
------------------------------------------------------------------------

pKeyOrIndex :: AT.Parser Key
pKeyOrIndex = pIndex <|> pKey

pKey = Key <$> AT.takeWhile1 (AT.notInClass " .[:")

pIndex = Index <$> AT.decimal <* AT.char ']'

evalToValues :: Text -> [[Key]] -> Value -> [Value]
evalToValues arrayDelim ks v = map (evalToValue arrayDelim v) ks

type ArrayDelimiter = Text

evalToValue :: ArrayDelimiter -> Value -> [Key] -> Value
evalToValue d v k = evalKeyPath d k v

evalToText :: ArrayDelimiter -> [Key] -> Value -> Text
evalToText d k v = valToText $ evalKeyPath d k v

-- evaluates the a JS key path against a Value context to a leaf Value
evalKeyPath :: ArrayDelimiter -> [Key] -> Value -> Value
evalKeyPath d [] x@(String _) = x
evalKeyPath d [] x@Null = x
evalKeyPath d [] x@(Number _) = x
evalKeyPath d [] x@(Bool _) = x
evalKeyPath d [] x@(Object _) = x
evalKeyPath d [] x@(Array v) | V.null v = Null
evalKeyPath d [] x@(Array v) = 
          let vs = V.toList v
              xs = intersperse d $ map (evalToText d []) vs
          in String . mconcat $ xs
evalKeyPath d (Key key:ks) (Object s) = 
    case (HM.lookup key s) of
        Just x          -> evalKeyPath d ks x
        Nothing -> Null
evalKeyPath d (Index idx:ks) (Array v) = 
      let e = (V.!?) v idx
      in case e of 
        Just e' -> evalKeyPath d ks e'
        Nothing -> Null
-- traverse array elements with additional keys
evalKeyPath d ks@(Key key:_) (Array v) | V.null v = Null
evalKeyPath d ks@(Key key:_) (Array v) = 
      let vs = V.toList v
      in String . mconcat . intersperse d $ map (evalToText d ks) vs
evalKeyPath _ ((Index _):_) _ = Null
evalKeyPath _ _ _ = Null

valToText :: Value -> Text
valToText (String x) = x
valToText Null = "null"
valToText (Bool True) = "t"
valToText (Bool False) = "f"
valToText (Number x) = 
    case floatingOrInteger x of
        Left float -> T.pack . show $ float
        Right int -> T.pack . show $ int
valToText (Object _) = "[Object]"



t = runTestTT tests

str1 = "1234567890"

tests = test [
    "no truncation, under limit " ~: str1         @=? truncateStr 10 str1
  , "truncation"                  ~: "12345..."   @=? truncateStr 8 str1
  , "John Smith bytelength"       ~: 10           @=? bytelen  "John Smith"
  , "John Smith"                  ~: "John Smith" @=? truncateStr 10 "John Smith"
  ]

