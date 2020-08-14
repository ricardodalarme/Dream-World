Attribute VB_Name = "modGlobals"
Option Explicit

' Used for loops
Public GiveNPCHPTimer As Long
Public LastUpdatePlayerVitals As Long

' Used for logging
Public ServerLog As Boolean

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean
