VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFastCollection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =======================================================
'
' vkUserControlsXP
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' Some graphical UserControls for your VB application.
'
' Copyright © 2006-2007 by Alain Descotes.
'
' vkUserControlsXP is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' vkUserControlsXP is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
' =======================================================

Option Explicit


'=======================================================
'APIS
'=======================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'=======================================================
'variables privées
'=======================================================
Private lCount As Long  'nombre d'Items
Private lUbound As Long 'taille du tableau
Private lPtr() As Long  'contient la liste des pointeurs vers les objets

Private Const USTEP As Long = 1000  'variable arbitraire pour le redimensionnement
'du tableau. Plus c'est grand, plus çà gaspille en mémoire (potentiellement)
'mais plus c'est rapide.



'=======================================================
'PUBLIC FUNCTIONS & PROPERTIES
'=======================================================

'=======================================================
'ajoute un item
'=======================================================
Public Sub Add(Obj As Object, Optional ByVal Index As Long = -1)
Attribute Add.VB_Description = "Add an item to collection"
Dim Tmp As Object
    
    'récupère l'objet passé en paramètre
    Set Tmp = Obj
    
    'on teste si c'est nécessaire de redimensionner le tableau ou pas
    If lCount = lUbound Then
    
        'alors la taille du tableau est trop petite avec le nouvel Item
        'donc on redimensionne ==> change le lUbound
        lUbound = lUbound + USTEP
        
        'on redimensionne avec Preserve avec la nouvelle taille
        ReDim Preserve lPtr(lUbound)
        
    End If
        
    'incrémente le nombre d'objets
    lCount = lCount + 1
    
    'si l'index est incorrect ==> alors c'est à la suite
    If Index <= 0 Or Index > lCount Then

        'recupère le pointeur sur l'objet et le stocke dans le tableau
        lPtr(lCount) = ObjPtr(Obj)

    Else
    
        'alors on insère dans la liste à l'Index voulu

        'on copie la partie du tableau APRES index en décalant d'une case
        'le tableau, grâce à Copymemory
        Call CopyMemory(lPtr(Index + 1), lPtr(Index), (lCount - Index) * 4)
    
        'on insère le nouvel item au milieu du tableau
        lPtr(Index) = ObjPtr(Obj)
                
    End If
    
    '# On supprime le lien qui existe (ObjPtr) entre Tmp et notre objet.
    '# Ainsi, la référence que nous venons de placer dans le tableau lPtr restera valide une fois la procédure
    '# terminée (une fois Tmp en dehors de la portée) : nous avons subtilisée une référence vers notre objet.
    Call CopyMemory(Tmp, 0&, 4&)
    
End Sub

'=======================================================
'enlève un item
'=======================================================
Public Sub Remove(Index As Long)
Attribute Remove.VB_Description = "Remove an item"
Dim Obj As Object
    
    'si l'index est correct, on supprime
    If Index > 0 And Index <= lCount Then

        '//libère l'objet correspondant à l'index à virer
        'copie l'Item dans la variable temporaire grâce à son adresse mémoire
        Call CopyMemory(Obj, lPtr(Index), 4)
        'destroy l'Item par Set=Nothing
        Set Obj = Nothing
    
        'on décale de 1 vers la gauche le tableau
        Call CopyMemory(lPtr(Index), lPtr(Index + 1), (lCount - Index) * 4)
    
        'change le nombre d'Items
        lCount = lCount - 1
        
        'redimensionne le tableau SI NECESSAIRE
        If lUbound - USTEP > lCount Then
            
            'on enlève USTEP items car les USTEP derniers ne servent plus
            lUbound = lUbound - USTEP
            
            'alors on redimensionne le tableau en virant USTEP éléments
            ReDim Preserve lPtr(lUbound)
        
        End If
        
    End If
    
End Sub

'=======================================================
'clear la collection
'=======================================================
Public Sub Clear()
Attribute Clear.VB_Description = "Clear the collection"
Dim ptrFirst As Long
Dim x As Long
Dim Obj As Object
    
    'que si des Items présents
    If lCount Then
        
        'supprime tous les objets de la liste.
        'ne parcours pas toute la liste : ne récupère que le pointeur du
        'premier item de la liste et destroy les items situés à chaque pointeur
        'Chaque élément du tableau est décalé de 4 bytes (Long) en mémoire
        
        'récupère le pointeur du premier objet
        ptrFirst = VarPtr(lPtr(1))
        
        For x = ptrFirst To ptrFirst + (lCount - 1) * 4 Step 4
            
            'on récupère l'objet dans la variable temp Obj
            Call CopyMemory(Obj, ByVal x, 4&) '4 <==> long
            
            'destroy l'item par Set=Nothing
            Set Obj = Nothing
            
        Next x
        
        'supprime le tableau et initialise les valeurs
        ReDim lPtr(0)
        lCount = 0: lUbound = 0
    
    End If
 
End Sub

'=======================================================
'récupère l'item pointé par l'index
'=======================================================
Public Function Item(Index As Long) As Object
Attribute Item.VB_Description = "Get an item"
Attribute Item.VB_UserMemId = 0
Dim Obj As Object
    
    'si l'index est valide
    If Index > 0 And Index <= lCount Then
        
        'on a besoin de récupérer notre objet dans une variable temporaire
        Call CopyMemory(Obj, lPtr(Index), 4)      '4 octets
        
        'récupère l'objet en résultat de la function
        Set Item = Obj
        
        'détruit l'objet temporaire maintenant inutile
        Call CopyMemory(Obj, 0&, 4)
        
    End If

End Function

'=======================================================
'récupère le Count de la collection
'=======================================================
Public Property Get Count() As Long: Count = lCount: End Property
Attribute Count.VB_Description = "Get the number of items in collection"





'=======================================================
'SUB DE LA CLASSE
'=======================================================

'=======================================================
'initialisation et destruction
'=======================================================
Private Sub Class_Initialize()
    lCount = 0: lUbound = 0: ReDim lPtr(0)
End Sub
Private Sub Class_Terminate()
    Call Clear
    Erase lPtr  'supprime complètement le tableau
End Sub
