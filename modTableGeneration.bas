Attribute VB_Name = "modTableGeneration"
Option Compare Database

'********************************************************************************************************
'   RC Table Generation - v0.1                                                                          *
'   Generates tables in the current DB as laid out in tblTABLE_SCHEMA                                   *
'   Fields :                                                                                            *
'               [TABLE_NAME]                                                                            *
'               [FIELD_NAME]                                                                            *
'               [FIELD_TYPE] "INTEGER", "STRING", "DECIMAL", "DATE", "YES/NO", "CURRENCY", "PERCENT"    *
'               [FIELD_DESC]                                                                            *
'               [PRIMARY_KEY]                                                                           *
'               [FOREIGN_KEY]                                                                           *
'               [INDEXED]                                                                               *
'               [REQUIRED]                                                                              *
'                                                                                                       *
'   Published 4/10/2014         Author : Bill Normandin                                                 *
'   Language : Visual Basic for Applications                                                            *
'********************************************************************************************************

Public Sub GenerateTables()
On Error GoTo Err_Handler:

    Dim fld As Field, fldName As String, tblName As String, lastTblName As String, tdf As DAO.TableDef, fldType As Variant, AutoIncr As Boolean, rq As Boolean
    Dim fldDesc As String, pkFlag As Boolean, fkFlag As Boolean, idxFlag As Boolean, flg As Boolean
    Dim rs As DAO.Recordset, db As DAO.Database, idx As DAO.Index
    
    Debug.Print "RC Property Master Backend - " & RC_GetVariable("App") & " version " & RC_GetVariable("Version")
    Debug.Print "Copyright 2014; Authored by Bill Normandin"
    Debug.Print "-------------------------------------------------------------------------------------------------"
    Debug.Print "Beginning Table Generation Routine : " & Now()
    Debug.Print "   ..."
    
    Debug.Print "   Opening DB and tblTABLE_SCHEMA..."
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblTABLE_SCHEMA ORDER BY FIELD_ID", dbOpenDynaset)
    lastTblName = "none"
    
    Debug.Print "   Success... Entering Table Generation Loop..."
    With rs
    
        Do Until .EOF

            tblName = ![TABLE_NAME]
            fldName = ![FIELD_NAME]
            fldType = ![FIELD_TYPE]
            fldDesc = ![FIELD_DESC]
            pkFlag = ![PRIMARY_KEY]
            fkFlag = ![FOREIGN_KEY]
            idxFlag = ![Indexed]
            rq = ![Required]
        
            If tblName <> lastTblName Then  'Check for a new table, close out the old table and generate a new TableDef
            
                If lastTblName <> "none" Then
                
                    Debug.Print "   Closing table " & lastTblName
                    db.TableDefs.Append tdf
                    db.TableDefs.Refresh
                    Set tdf = Nothing
                    
                End If
                
                Set tdf = db.CreateTableDef(tblName)
                Debug.Print "   Created table " & tblName & "..."
                
            End If
            
            If DataType(fldType) = 4 And pkFlag = True Then
            
                AutoIncr = True
                
            End If
            
            Debug.Print "   Field Name : " & fldName
            Debug.Print "   Attributes :"
            Debug.Print "       Type - " & DataType(fldType) & " ( " & fldType & " )"
            If idxFlag Then Debug.Print "       Indexed"
            If pkFlag Then Debug.Print "       Primary Key"
            If AutoIncr Then Debug.Print "     Auto-Incremented Field"
            If rq Then Debug.Print "        Required"
            Debug.Print "       Description : " & fldDesc
            
            Debug.Print "   Creating field..."
            Set fld = tdf.CreateField(fldName, DataType(fldType))
            FieldAttributes fld, AutoIncr, rq
            
            tdf.Fields.Append fld
            
            Set fld = Nothing
            
            If idxFlag Then
            
                Debug.Print "   Creating Index..."
                Set idx = tdf.CreateIndex(fldName)
                Set fld = idx.CreateField(fldName)
                idx.Fields.Append fld
                
                If pkFlag Then
                    idx.Primary = True
                    Debug.Print "   Set Primary Key..."
                End If
                
                tdf.Indexes.Append idx
                Set fld = Nothing
                
            End If
        
            Debug.Print "   Done!"
            lastTblName = tblName
            AutoIncr = False
            .MoveNext
            
        Loop
        
    End With
    
    'Append the last table
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
    
GenerateTables_Exit:

    On Error Resume Next
    
    Application.RefreshDatabaseWindow
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set fld = Nothing
    Set tdf = Nothing
    Exit Sub
            
Err_Handler:

    RC_Log "Table Generation Error", "Module GenerateTable raised an exception.", "Error", Err.Number, Err.Description, True
    Resume GenerateTables_Exit

End Sub

Public Sub FieldAttributes(fld As Field, AutoIncr As Boolean, rq As Boolean)

    With fld
        If rq Then .Required = True
        If AutoIncr Then .Attributes = dbAutoIncrField
    End With

End Sub

Public Function DataType(ByVal str As String) As Long

    Dim sOut As Long

    Select Case str
    
        Case "INTEGER"
        
            sOut = dbLong
            GoTo DataType_Exit
        
        Case "STRING"
        
            sOut = dbText
            GoTo DataType_Exit
        
        Case "YES/NO"
        
            sOut = dbBoolean
            GoTo DataType_Exit
        
        Case "CURRENCY"
        
            sOut = dbCurrency
            GoTo DataType_Exit
            
        Case "DATE"
        
            sOut = dbDate
            GoTo DataType_Exit
            
        Case "DECIMAL"
        
            sOut = dbDouble
            GoTo DataType_Exit
            
        Case "PERCENT"
        
            sOut = dbDouble
            GoTo DataType_Exit
            
        Case "DATE/TIME"
        
            sOut = dbDate
            GoTo DataType_Exit
        
    End Select
    
DataType_Exit:

    DataType = sOut

End Function

