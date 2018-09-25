VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------
'|    _       _       _             |
'|   | |     (_)     | |            |
'|   | |      _  ___ | |_           |
'|   | |     | |/ __|| __|          |
'|   | |____ | |\__ \| |_           |
'|   |______||_||___/ \__|          |
'|    _____  _                      |
'|   |_   _|| |                     |
'|     | |  | |_  ___  _ __ ___     |
'|     | |  | __|/ _ \| '_ ` _ \    |
'|    _| |_ | |_|  __/| | | | | |   |
'|   |_____| \__|\___||_| |_| |_|   |
'|                                  |
'|----------------------------------|
'| Items used by Linked List        |
'|----------------------------------|
'| VBA does not allow recursion in  |
'|  typedefs so small classes like  |
'|  this need to be made            |
'|----------------------------------|
'| By Ryan Hoover                   |
'------------------------------------

Public Value As Variant
Public NextItem As ClsListItem
