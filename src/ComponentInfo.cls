VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Doc As ModelDoc2
Public Conf As String
Public PropertyDesignation As String 'Обозначение
Public PropertyName As String 'Наименование
Public PropertyBlank As String 'Заготовка
Public BaseDesignation As String 'Базовое Обозначение (без исполнения)
Public Drawings As Dictionary 'of DrawingRecord
Public Parent As String
