
function InitializeSheet()
{
  //Cell information
  var numberOfPlayerCell = "C14" //Where to pull number of players from
  var cardNumberColumn = "A" //Column the card number displayed in
  var numberRowStart = 3 //Row where numbering should start
  var trumpColumn = "B" //Column where trump is displayed
  var scoreSummationStart = "C" //First column under first player
  var labelRow = 2 //Row where bid, made, and score labels are
  
  
  var initializationSheet = SpreadsheetApp.getActive().getSheetByName("Instructions")
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Scorecard")
  
  //Getting Player names and count
  var playerRange = initializationSheet.getRange("B3:B11")
  var playerRangeArray = playerRange.getValues()
  var players = []
  for (var i = 0; i < playerRangeArray.length; i++)
  {
    if(playerRangeArray[i][0] != "Enter Player Here" && playerRangeArray[i][0] != "") {players.push(playerRangeArray[i][0])}
  }
  
  var numberOfPlayers = players.length
  
  //Clears anything expect for permamanent labels
  spreadsheet.getRange("A3:AC30").clearContent()
  spreadsheet.getRange("C1:AC2").clearContent()
  
  //This loop puts in the trump and card number round
  var trumpCycle = ["Clubs", "Diamonds", "Hearts", "Spades", "No Trump"]
  var numberOfRounds = Math.floor(52/numberOfPlayers)
  var currentRound = 0
  for (var i = numberOfRounds; i > 0; i--)
  {
    var cardRange = cardNumberColumn + (numberRowStart+currentRound) //Increments down cardNumberColumn
    spreadsheet.getRange(cardRange).setValue(i)
    
    var trumpRange = trumpColumn + (numberRowStart+currentRound) //Increments down trumpColumn
    spreadsheet.getRange(trumpRange).setValue(trumpCycle[currentRound%5])
     
    currentRound++
  }
  
  //This loop handles setting up a score cell for each player and filling in the 'hidden' score column for each player
  var scoreRow = numberRowStart+currentRound+1
  var playerNameCount = 0
  spreadsheet.getRange(trumpColumn + scoreRow).setValue("Scores")
  for (var i = 3; i <= numberOfPlayers*3; i+=3)
  {
    //Setting each players name
    spreadsheet.getRange(columnToLetter(i)+1).setValue(players[playerNameCount++])
    //Giving each player a score cell that sums up the row 2 to the right of the cell
    spreadsheet.getRange(scoreRow, i).setValue("=SUM(" + columnToLetter(i+2) + numberRowStart + ":" + columnToLetter(i+2) + (scoreRow-1) + ")")
    
    //For every player, loop down their score column and fill it with a dynamic sum function with nested if statements its ugly I know.
    for(var j = numberRowStart; j <= (scoreRow-2); j++)
    {
      var currentColumn = columnToLetter(i+2)
      var columnRange = spreadsheet.getRange(j, i+2)
      var ifString = "=IF("+ columnToLetter(i+1)+j + "=0, 0, IF(" + columnToLetter(i) + j + "=" + columnToLetter(i+1)+j + ", $A" + j + "+" + columnToLetter(i+1)+j + ", $A" + j + "-ABS(" + columnToLetter(i+1)+j+ "-" + columnToLetter(i) + j + ")))"
      columnRange.setValue(ifString)
    }
    for(var k = 0; k < 3; k++)
    {
      var labelCycle = ["Bid", "Taken", "Score"]
      spreadsheet.getRange(labelRow, i+k).setValue(labelCycle[k%3])
    }
    
    SpreadsheetApp.setActiveSheet(spreadsheet)
  }

  
  
  
  
  //Simple function to convert column indexes to a1 lettering I found on the internet
  function columnToLetter(column)
  {
    var temp, letter = '';
    while (column > 0)
    {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }
  
}