Attribute VB_Name = "GLOBALS"

' === GLOBALS.bas ===
Option Explicit
Public Const TO_INJECT_PATH As String = "vbaCodebase\toInject\"

' === Feuille principale de production
Public PRODUCTION_WS As Worksheet

' === Adresse de la zone maximale du rouleau
Public Const MAX_ROLL_ADDR As String = "AJ68:BD167"

' === Adresse de la cellule contenant la longueur cible du rouleau
Public Const TARGET_LENGTH_ADDR As String = "BH71"
' === Objet Range pointant vers cette cellule
Public TARGET_LENGTH_CELL As Range

