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
Public PropertyDesignation As String '�����������
Public PropertyName As String '������������
Public PropertyBlank As String '���������
Public BaseDesignation As String '������� ����������� (��� ����������)
Public Drawings As Dictionary 'of DrawingRecord
Public Parent As String
