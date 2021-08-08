VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CreateRelationX()

    Dim dbsNorthwind As Database
    Dim tdfEmployees As TableDef
    Dim tdfNew As TableDef
    Dim idxNew As Index
    Dim relNew As Relation
    Dim idxLoop As Index

    Set dbsNorthwind = OpenDatabase("Northwind.mdb")

    With dbsNorthwind
        ' Add new field to Employees table.
        Set tdfEmployees = .TableDefs!Employees
        tdfEmployees.Fields.Append _
            tdfEmployees.CreateField("DeptID", dbInteger, 2)

        ' Create new Departments table.
        Set tdfNew = .CreateTableDef("Departments")

        With tdfNew
            ' Create and append Field objects to Fields
            ' collection of the new TableDef object.
            .Fields.Append .CreateField("DeptID", dbInteger, 2)
            .Fields.Append .CreateField("DeptName", dbText, 20)

            ' Create Index object for Departments table.
            Set idxNew = .CreateIndex("DeptIDIndex")
            ' Create and append Field object to Fields
            ' collection of the new Index object.
            idxNew.Fields.Append idxNew.CreateField("DeptID")
            ' The index in the primary table must be Unique in
            ' order to be part of a Relation.
            idxNew.Unique = True
            .Indexes.Append idxNew
        End With

        .TableDefs.Append tdfNew

        ' Create EmployeesDepartments Relation object, using
        ' the names of the two tables in the relation.
        Set relNew = .CreateRelation("EmployeesDepartments", _
            tdfNew.Name, tdfEmployees.Name, _
            dbRelationUpdateCascade)

        ' Create Field object for the Fields collection of the
        ' new Relation object. Set the Name and ForeignName
        ' properties based on the fields to be used for the
        ' relation.
        relNew.Fields.Append relNew.CreateField("DeptID")
        relNew.Fields!DeptID.ForeignName = "DeptID"
        .Relations.Append relNew

        ' Print report.
        Debug.Print "Properties of " & relNew.Name & _
            " Relation"
        Debug.Print "  Table = " & relNew.Table
        Debug.Print "  ForeignTable = " & _
            relNew.ForeignTable
        Debug.Print "Fields of " & relNew.Name & " Relation"

        With relNew.Fields!DeptID
            Debug.Print "  " & .Name
            Debug.Print "    Name = " & .Name
            Debug.Print "    ForeignName = " & .ForeignName
        End With

        Debug.Print "Indexes in " & tdfEmployees.Name & _
            " TableDef"
        For Each idxLoop In tdfEmployees.Indexes
            Debug.Print "  " & idxLoop.Name & _
                ", Foreign = " & idxLoop.Foreign
        Next idxLoop

        ' Delete new objects because this is a demonstration.
        .Relations.Delete relNew.Name
        .TableDefs.Delete tdfNew.Name
        tdfEmployees.Fields.Delete "DeptID"
        .Close
    End With

End Sub
 
Private Sub ShowRelations()

    Dim db1 As DAO.Database
    Dim relLoop As Relation
    Dim fld As Field
    Dim pwd As String
    pwd = "Enya"
    
    Set db1 = DBEngine.OpenDatabase(App.Path & "\Llamadas.mdb", True, False, ";Pwd=" & pwd)
    Debug.Print db1.Relations.Count
    For Each relLoop In db1.Relations
        Debug.Print "------------------------"
        Debug.Print "Relación "; relLoop.Name
        Debug.Print "         "; relLoop.Attributes
        Debug.Print "T/F      "; relLoop.Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade
        Debug.Print "Foreing  "; relLoop.ForeignTable
        Debug.Print "Table    "; relLoop.Table
        Debug.Print "Fields   "; relLoop.Fields.Count
        For Each fld In relLoop.Fields
            Debug.Print "Field    "; fld.Name
            Debug.Print "Foreing  "; fld.ForeignName
        Next
    Next
    
    db1.Close
    
End Sub

Private Sub Command1_Click()
    ShowRelations
End Sub

Private Sub MakeRelations()
    Dim db1 As DAO.Database
    Dim rel As Relation
    Dim fld As Field
    Dim pwd As String
    pwd = "Enya"
    
    Set db1 = DBEngine.OpenDatabase(App.Path & "\Llamadas.mdb", True, False, ";Pwd=" & pwd)
    
    'crea nueva relación usando los nombres de las tablas
    Set rel = db1.CreateRelation("GruposUsuariosxGrupo")

    With rel
        'establece atributos y tablas
        .Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade
        .Table = "Grupos"
        .ForeignTable = "UsuariosxGrupo"
        
        'campos de la tabla Grupos
        With rel
            .Fields.Append .CreateField("ProyectoCod")
            .Fields.Append .CreateField("CategoriaCod")
            .Fields.Append .CreateField("CCostoCod")
        End With
        
        'campos de la tabla externa UsuariosxGrupo
        For Each fld In rel.Fields
            fld.ForeignName = fld.Name
        Next
    End With
    'agrega rel a la colección relations de db1
    db1.Relations.Append rel
    
End Sub

Private Sub Command2_Click()
    Dim db1 As DAO.Database
    Dim rel As Relation
    Dim fld As Field
    Dim pwd As String
    pwd = "Enya"
    
    Set db1 = DBEngine.OpenDatabase(App.Path & "\Llamadas.mdb", True, False, ";Pwd=" & pwd)

    Dim i As Integer
    Do While db1.Relations.Count > 0
        db1.Relations.Delete db1.Relations(0).Name
    Loop
    db1.Close
End Sub

Private Sub Command3_Click()
    MakeRelations
End Sub
