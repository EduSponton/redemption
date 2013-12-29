VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    
Public Function CreateSelectQuery(ByVal sSelect As String, ByVal From As String, ByVal Where As String, ByVal GroupBy As String, OrderBy As String) As String
    Dim sSelectAux As String
    Dim sFrom As String
    Dim sWhere As String
    Dim sGroupBy As String
    Dim sOrderBy As String
    
    
    sSelectAux = "SELECT " & sSelect & " "
    sFrom = "FROM " & From & " "
    sWhere = "WHERE " & Where
    If Trim$(GroupBy) = "" Then
        sGroupBy = ""
    Else
        sGroupBy = "GROUP BY " & GroupBy
    End If
    
    
    If Trim$(OrderBy) = "" Then
        sOrderBy = ""
    Else
        sOrderBy = "ORDER BY " & OrderBy
    End If
    

    CreateSelectQuery = sSelectAux & " " & sFrom & " " & sWhere & " " & sGroupBy & " " & sOrderBy & ";"
End Function

Public Function CreateUpdateQuery(ByVal Table As String, ByVal campos As HashTableObject, ByVal Where As String) As String
    Dim sUpdate As String
    Dim sSet As String
    Dim sWhere As String
    
    Dim campoTMP As campo
    sUpdate = "UPDATE " & Table & " "
    sSet = "SET "
    sWhere = "WHERE " & Where
    
    
    With campos
        .MoveFirst
        While .MoveNext
            Set campoTMP = .GetCurrentValue
            If .IsLast = True Then
              '  Debug.Print .IsLast
              
                If campoTMP.getTipoDato = "string" Then
                    sSet = sSet & campoTMP.getNombre & "= '" & campoTMP.getValor & "'"
                Else
                    sSet = sSet & campoTMP.getNombre & "= " & campoTMP.getValor
                End If
            Else
                If campoTMP.getTipoDato = "string" Then
                    sSet = sSet & campoTMP.getNombre & "= '" & campoTMP.getValor & "', "
                Else
                    sSet = sSet & campoTMP.getNombre & "= " & campoTMP.getValor & ", "
                End If
            End If
        Wend
    End With
    
    CreateUpdateQuery = sUpdate & " " & sSet & " " & sWhere & ";"
        
End Function

Public Function CreateInsertQuery(ByVal Table As String, ByVal campos As HashTableObject) As String
    Dim sInsert As String
    Dim sColums As String
    Dim sValues As String
    Dim campoTMP As campo
    
    sInsert = "INSERT INTO " & Table
    sColums = "("
    sValues = "VALUES ("
    
    With campos
        .MoveFirst
        While .MoveNext
            Set campoTMP = .GetCurrentValue
            If .IsLast = True Then
                sColums = sColums & campoTMP.getNombre & ") "
                If campoTMP.getTipoDato = "string" Then
                    sValues = sValues & "'" & campoTMP.getValor & "')"
                Else
                    sValues = sValues & campoTMP.getValor & ")"
                End If
            Else
                sColums = sColums & campoTMP.getNombre & ", "
                If campoTMP.getTipoDato = "string" Then
                    sValues = sValues & "'" & campoTMP.getValor & "', "
                Else
                    sValues = sValues & campoTMP.getValor & ", "
                End If
                
            End If
            
            
        Wend
    End With
    CreateInsertQuery = sInsert & " " & sColums & " " & sValues & ";"
    
End Function

Public Function CreateDeleteQuery(ByVal Table As String, ByVal Where As String) As String
    Dim sDelete As String
    Dim sWhere As String
    
    sDelete = "DELETE FROM " & Table & " "
    sWhere = "WHERE " & Where
    
    CreateDeleteQuery = sDelete & " " & sWhere & ";"
    
End Function

Public Function WhereCreator(ByVal ConditionAndLogicalOperation As HashTable) As String

    Dim sWhere As String
    sWhere = ""
    With ConditionAndLogicalOperation
        .MoveFirst
        While .MoveNext
            If .IsLast = True Then
                sWhere = sWhere & " (" & .GetCurrentKey & ")"
            Else
                sWhere = sWhere & " (" & .GetCurrentKey & ") " & .GetCurrentValue
            End If
            
        Wend
    End With
    WhereCreator = sWhere
End Function


Public Function FromCreator(ByVal Tables As CollectionArr) As String
    Dim i As Integer
    Dim sFrom As String
    With Tables
        For i = .GetFirstIndex To .GetLastIndex
            If i = .GetLastIndex Then
                sFrom = sFrom & .Item(i)
            Else
                sFrom = sFrom & .Item(i) & ", "
            End If
        Next
    End With
    FromCreator = sFrom
End Function

Public Function FromWithAliasCreator(ByVal TablesAndAlias As HashTable) As String
    Dim sFrom As String
    
    With TablesAndAlias
        .MoveFirst
        While .MoveNext
            If .IsLast = True Then
                sFrom = sFrom & .GetCurrentKey & " " & .GetCurrentValue
            Else
                sFrom = sFrom & .GetCurrentKey & " " & .GetCurrentValue & ", "
            End If
        Wend
    End With
    FromWithAliasCreator = sFrom
End Function

Public Function GroupByCreator(ByVal Colums As CollectionArr) As String
    Dim i As Integer
    Dim sGroupBy As String
    With Colums
        For i = .GetFirstIndex To .GetLastIndex
            If i = .GetLastIndex Then
                sGroupBy = sGroupBy & .Item(i)
            Else
                sGroupBy = sGroupBy & .Item(i) & ", "
            End If
        Next
    End With
    GroupByCreator = sGroupBy
            
End Function

Public Function SelectCreator(ByVal Colums As CollectionArr) As String
    Dim i As Integer
    Dim sSelectCreator As String
    With Colums
        For i = .GetFirstIndex To .GetLastIndex
            If i = .GetLastIndex Then
                sSelectCreator = sSelectCreator & .Item(i)
            Else
                sSelectCreator = sSelectCreator & .Item(i) & ", "
            End If
        Next
    End With
    sSelectCreator = sSelectCreator
End Function

Public Function SelectWithAliasCreator(ByVal ColumsAndAlias As HashTable) As String
    Dim sSelect As String
    
    With ColumsAndAlias
        .MoveFirst
        While .MoveNext
            If .IsLast = True Then
                sSelect = sSelect & .GetCurrentKey & " as '" & .GetCurrentValue & "'"
            Else
                sSelect = sSelect & .GetCurrentKey & " as '" & .GetCurrentValue & "', "
            End If
        Wend
    End With
    SelectWithAliasCreator = sSelect
End Function
Public Function OrderByCreator(ByVal ColumsAndOrder As HashTable) As String
    Dim sOrderBy As String
    
    With ColumsAndOrder
        .MoveFirst
        While .MoveNext
            If .IsLast = True Then
                sOrderBy = sOrderBy & .GetCurrentKey & " " & .GetCurrentValue
            Else
                sOrderBy = sOrderBy & .GetCurrentKey & " " & .GetCurrentValue & ", "
            End If
        Wend
    End With
    OrderByCreator = sOrderBy
        
End Function

Public Function CaseCreator(ByVal ColumnValue As String, ByVal WhenValuesAndThenValues As HashTable) As String
    Dim sColumnValue As String
    Dim sWhenValuesAndThenValues As String
    
    sColumnValue = "CASE (" & ColumnValue & ") "
    sWhenValuesAndThenValues = ""
    
    With WhenValuesAndThenValues
        .MoveFirst
        While .MoveNext
            If .IsLast = True Then
                sWhenValuesAndThenValues = sWhenValuesAndThenValues & "WHEN " & .GetCurrentKey & " THEN " & .GetCurrentValue & " END"
            Else
                sWhenValuesAndThenValues = sWhenValuesAndThenValues & "WHEN " & .GetCurrentKey & " THEN " & .GetCurrentValue & " "
            End If
        Wend
    End With
    
    CaseCreator = sColumnValue & " " & sWhenValuesAndThenValues
End Function


Public Function IFCreator(ByVal Condition As String, ByVal TrueValue As String, ByVal FalseValue As String) As String
    IFCreator = "IF ((" & Condition & ")," & TrueValue & "," & FalseValue & ")"
End Function

Public Function BetweenCreator(ByVal Column As String, ByVal BaseValue As String, ByVal TopValue As String) As String
    BetweenCreator = Column & " BETWEEN " & BaseValue & " AND " & TopValue
End Function

Public Function LikeCreator(ByVal Column As String, ByVal StringSearch As String) As String
    LikeCreator = Column & " LIKE '" & StringSearch & "'"

End Function
