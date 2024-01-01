

function searchRowIndex(text, sheet, targetColumnIndex = 1, startRowIndex = 1, endRowIndex = null)
{
  if(null == endRowIndex || sheet.getLastRow() < endRowIndex)
  {
    endRowIndex = sheet.getLastRow()
  }
  var ret = -1;
  for(var row = startRowIndex; row <= endRowIndex; row++)
  {
    var cell = sheet.getRange(row, targetColumnIndex);
    var value = (cell.isPartOfMerge() ? cell.getMergedRanges()[0].getCell(1,1) : cell).getDisplayValue();
//console.log("value: " + value + ", text: " + text);
    if(value == text)
    {
      ret = row;
      break;
    }
  }
  return ret;
}


function searchColumnIndex(text, sheet, targetRowIndex = 1, startColumnIndex = 1, endColumnIndex = null)
{
  if(null == endColumnIndex || sheet.getLastColumn() < endColumnIndex)
  {
    endColumnIndex = sheet.getLastColumn()
  }
  var ret = -1;
  for(var col = startColumnIndex; col <= endColumnIndex; col++)
  {
    var cell = sheet.getRange(targetRowIndex, col);
    var value = (cell.isPartOfMerge() ? cell.getMergedRanges()[0].getCell(1,1) : cell).getDisplayValue();
    if(value == text)
    {
      ret = col;
      break;
    }
  }
  return ret;
}


function searchNextRowIndex(text, sheet, targetColumnIndex = 1, startRowIndex = 1, endRowIndex = null)
{
  if(null == endRowIndex || sheet.getLastRow() < endRowIndex)
  {
    endRowIndex = sheet.getLastRow()
  }
  var row = searchRowIndex(text, sheet, targetColumnIndex, startRowIndex, endRowIndex);
  var cell = sheet.getRange(row, targetColumnIndex);
  return (cell.isPartOfMerge() ? cell.getMergedRanges()[0].getLastRow() : row) + 1;
}

