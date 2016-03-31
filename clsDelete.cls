VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum ReMasterDelete
    reDelMasterCurency = 0
    reDelMasterGudang = 1
    reDelMasterKelompok = 2
    reDelMasterBarang = 3
    reDelMasterRegional = 4
    reDelMasterBayar = 5
    reDelMasterTransport = 6
    reDelMasterGudangCust = 7
    reDelMasterBank = 8
    reDelMasterPartner = 9
    reDelMasterAktiva = 10
    reDelMasterResources = 11
    reDelMasterStage = 12
    reDelMasterAccount = 13
End Enum

Public Function CekDelete(ByVal ParameterMaster As String, ByVal TypeDeleteData As ReMasterDelete) As Boolean
Dim rcDelete As New DBQuick
Select Case TypeDeleteData
       Case reDelMasterAccount:
            'TABEL JOURNAL
            rcDelete.DBOpen "SELECT [table journal].NoAccount FROM [table journal] INNER JOIN GLAccount ON [table journal].NoAccount = GLAccount.NoAccount " & _
            " GROUP BY [table journal].NoAccount HAVING ([table journal].NoAccount = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
            'DETIL JOURNAL
            rcDelete.DBOpen "SELECT GLAccount.NoAccount FROM GLAccount INNER JOIN [Detail Journal] ON GLAccount.NoAccount = [Detail Journal].NoAccount " & _
            " GROUP BY GLAccount.NoAccount HAVING (GLAccount.NoAccount = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            'INVENTORY
            rcDelete.DBOpen "SELECT GLAccount.NoAccount FROM GLAccount INNER JOIN Inventory ON GLAccount.NoAccount = Inventory.NoAccount " & _
            " GROUP BY GLAccount.NoAccount HAVING (GLAccount.NoAccount = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            'INVENTORY GROUP
            rcDelete.DBOpen "SELECT GLAccount.NoAccount FROM GLAccount INNER JOIN [Inventory Group] ON GLAccount.NoAccount = [Inventory Group].NoAccount " & _
            " GROUP BY GLAccount.NoAccount HAVING (GLAccount.NoAccount = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            'tabel rl
            'tabel pembantu
            
       Case reDelMasterCurency:
            rcDelete.DBOpen "SELECT  [PO Order].CurrID FROM  [PO Order] INNER JOIN [Currency Setup] ON [PO Order].CurrID = [Currency Setup].CurrID GROUP BY [PO Order].CurrID HAVING      ([PO Order].CurrID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     [Bank Partner].Currency FROM         [Currency Setup] RIGHT OUTER JOIN  [Bank Partner] ON [Currency Setup].CurrID = [Bank Partner].Currency GROUP BY [Bank Partner].Currency HAVING      ([Bank Partner].Currency = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterGudang:
            rcDelete.DBOpen "SELECT Inventory.WareHouse FROM WareHouse INNER JOIN Inventory ON WareHouse.WareHouse = Inventory.WareHouse GROUP BY Inventory.WareHouse HAVING      (Inventory.WareHouse = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     Inventory.WareHouse FROM         WareHouse INNER JOIN Inventory ON WareHouse.WareHouse = Inventory.WareHouse INNER JOIN [Detail TransData] ON Inventory.NoItem = [Detail TransData].NoItem GROUP BY Inventory.WareHouse HAVING      (Inventory.WareHouse = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     WareHouse.WareHouse FROM         TransData INNER JOIN                       WareHouse ON TransData.WareHouse = WareHouse.WareHouse WHERE     (WareHouse.WareHouse = N'" & ParameterMaster & "')", CNN
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
       Case reDelMasterKelompok:
            rcDelete.DBOpen "SELECT     [Inventory Group].NoGroup FROM         Inventory INNER JOIN [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup GROUP BY [Inventory Group].NoGroup HAVING      ([Inventory Group].NoGroup = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
       Case reDelMasterBarang:
            rcDelete.DBOpen "SELECT [Detail PO].NoItem FROM Inventory INNER JOIN [Detail PO] ON Inventory.NoItem = [Detail PO].NoItem GROUP BY [Detail PO].NoItem HAVING      ([Detail PO].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     [Detail TransData].NoItem FROM         Inventory INNER JOIN [Detail TransData] ON Inventory.NoItem = [Detail TransData].NoItem GROUP BY [Detail TransData].NoItem HAVING      ([Detail TransData].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT [Detail PO].NoItem FROM Inventory INNER JOIN [Detail PO] ON Inventory.NoItem = [Detail PO].NoItem GROUP BY [Detail PO].NoItem HAVING      ([Detail PO].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     [Raw Materials].NoItem FROM         Inventory INNER JOIN                       [Raw Materials] ON Inventory.NoItem = [Raw Materials].NoItem GROUP BY [Raw Materials].NoItem HAVING      ([Raw Materials].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     [Inventory Tabel].NoItem FROM         Inventory INNER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem GROUP BY [Inventory Tabel].NoItem HAVING      ([Inventory Tabel].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterRegional:
            rcDelete.DBOpen "SELECT     [Gudang Customer].RG FROM         Regional INNER JOIN [Gudang Customer] ON Regional.RG = [Gudang Customer].RG GROUP BY [Gudang Customer].RG HAVING      ([Gudang Customer].RG = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     PartnerDB.City FROM         Regional INNER JOIN                      PartnerDB ON Regional.[RG Name] = PartnerDB.City", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterBayar:
            rcDelete.DBOpen "SELECT     [Type Bayar].TypeFreight FROM         [Type Bayar] INNER JOIN                       [PO Order] ON [Type Bayar].TypeLoco = [PO Order].TypeLoco", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterTransport:
            rcDelete.DBOpen "SELECT [PO Order].ID FROM         [PO Order] LEFT OUTER JOIN                       Transport ON [PO Order].ID = Transport.ID GROUP BY [PO Order].ID HAVING      ([PO Order].ID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            'MessageBox rcDelete.PrepareQuery
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT TransData.ID FROM         Transport INNER JOIN                       TransData ON Transport.ID = TransData.ID GROUP BY TransData.ID HAVING      (TransData.ID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterGudangCust:
            rcDelete.DBOpen "SELECT     TransData.[GDG ID] FROM         TransData INNER JOIN                       [Gudang Customer] ON TransData.[GDG ID] = [Gudang Customer].[GDG ID] GROUP BY TransData.[GDG ID] HAVING      (TransData.[GDG ID] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterBank:
            rcDelete.DBOpen "SELECT     [Bank Partner].Account FROM         [PO Order] LEFT OUTER JOIN  [Bank Partner] ON [PO Order].Account = [Bank Partner].Account GROUP BY [Bank Partner].Account HAVING      ([Bank Partner].Account = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterPartner:
            rcDelete.DBOpen "SELECT     PartnerID FROM         [PO Order] GROUP BY PartnerID HAVING      (PartnerID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     PartnerId FROM         TransData GROUP BY PartnerId HAVING      (PartnerId = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
       Case reDelMasterAktiva:
            rcDelete.DBOpen "SELECT     [No Aktiva] FROM         [DTR Aktiva Tetap] WHERE     ([No Aktiva] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT     [Doc Reff] FROM         [Detail Journal] WHERE     ([Doc Reff] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
       Case reDelMasterResources:
            rcDelete.DBOpen "SELECT TypeID FROM [Resources Table] GROUP BY TypeID HAVING      (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT TypeID FROM wcenter_resources GROUP BY TypeID HAVING  (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT     TypeID FROM         [Manufacture Stage] GROUP BY TypeID HAVING      (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekDelete = True
               Exit Function
            End If
            
End Select
rcDelete.CloseDB
End Function

Public Function CekCascadeDeleteTable(ByVal ParameterMaster As String, ByVal TypeDeleteData As ReMasterDelete) As String
Dim rcDelete As New DBQuick
Select Case TypeDeleteData
       Case reDelMasterCurency:
            rcDelete.DBOpen "SELECT  [PO Order].CurrID FROM  [PO Order] INNER JOIN [Currency Setup] ON [PO Order].CurrID = [Currency Setup].CurrID GROUP BY [PO Order].CurrID HAVING      ([PO Order].CurrID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Purchasing / Sales"
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     [Bank Partner].Currency FROM         [Currency Setup] RIGHT OUTER JOIN  [Bank Partner] ON [Currency Setup].CurrID = [Bank Partner].Currency GROUP BY [Bank Partner].Currency HAVING      ([Bank Partner].Currency = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Bank Partner"
               Exit Function
            End If
       Case reDelMasterGudang:
            rcDelete.DBOpen "SELECT Inventory.WareHouse FROM WareHouse INNER JOIN Inventory ON WareHouse.WareHouse = Inventory.WareHouse GROUP BY Inventory.WareHouse HAVING      (Inventory.WareHouse = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = True
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     Inventory.WareHouse FROM         WareHouse INNER JOIN Inventory ON WareHouse.WareHouse = Inventory.WareHouse INNER JOIN [Detail TransData] ON Inventory.NoItem = [Detail TransData].NoItem GROUP BY Inventory.WareHouse HAVING      (Inventory.WareHouse = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Master Barang"
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     WareHouse.WareHouse FROM         TransData INNER JOIN                       WareHouse ON TransData.WareHouse = WareHouse.WareHouse WHERE     (WareHouse.WareHouse = N'" & ParameterMaster & "')", CNN
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Master Gudang"
               Exit Function
            End If
            
       Case reDelMasterKelompok:
            rcDelete.DBOpen "SELECT     [Inventory Group].NoGroup FROM         Inventory INNER JOIN [Inventory Group] ON Inventory.NoGroup = [Inventory Group].NoGroup GROUP BY [Inventory Group].NoGroup HAVING      ([Inventory Group].NoGroup = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Master Kelompok Barang"
               Exit Function
            End If
            
       Case reDelMasterBarang:
            'PURCHASE/SALES
            rcDelete.DBOpen "SELECT [Detail PO].NoItem FROM Inventory INNER JOIN [Detail PO] ON Inventory.NoItem = [Detail PO].NoItem GROUP BY [Detail PO].NoItem HAVING      ([Detail PO].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Purchasing / Sales"
               Exit Function
            End If
            'FIFO
            rcDelete.DBOpen "SELECT     [Detail TransData].NoItem FROM         Inventory INNER JOIN [Detail TransData] ON Inventory.NoItem = [Detail TransData].NoItem GROUP BY [Detail TransData].NoItem HAVING      ([Detail TransData].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Voucher"
               Exit Function
            End If
            'BOM COMPONEN DETAIL
            rcDelete.DBOpen "SELECT [BOM Component Detail].NoItem FROM Inventory INNER JOIN [BOM Component Detail] ON Inventory.NoItem = [BOM Component Detail].NoItem GROUP BY [BOM Component Detail].NoItem HAVING      ([BOM Component Detail].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Bill of Material"
               Exit Function
            End If
            'TRANSKASI JOURNAL
            rcDelete.DBOpen "SELECT     [Inventory Tabel].NoItem FROM         Inventory INNER JOIN [Inventory Tabel] ON Inventory.NoItem = [Inventory Tabel].NoItem GROUP BY [Inventory Tabel].NoItem HAVING      ([Inventory Tabel].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Journal"
               Exit Function
            End If
'            BOM STAGE DETAIL
            rcDelete.DBOpen "SELECT     [BOM Stage Detail].NoItem FROM         Inventory INNER JOIN [BOM Stage Detail] ON Inventory.NoItem = [BOM Stage Detail].NoItem GROUP BY [BOM Stage Detail].NoItem HAVING      ([BOM Stage Detail].NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "BOM Routing"
               Exit Function
            End If
'            backflush_output
            rcDelete.DBOpen "SELECT     backflush_output.NoItem FROM         Inventory INNER JOIN backflush_output ON Inventory.NoItem = backflush_output.NoItem GROUP BY backflush_output.NoItem HAVING      (backflush_output.NoItem = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Backflush Finished Goods"
               Exit Function
            End If
       Case reDelMasterRegional:
            rcDelete.DBOpen "SELECT     [Gudang Customer].RG FROM         Regional INNER JOIN [Gudang Customer] ON Regional.RG = [Gudang Customer].RG GROUP BY [Gudang Customer].RG HAVING      ([Gudang Customer].RG = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Gudang Customer"
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     PartnerDB.City FROM         Regional INNER JOIN                      PartnerDB ON Regional.[RG Name] = PartnerDB.City", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Master Kota Partner"
               Exit Function
            End If
       Case reDelMasterBayar:
            rcDelete.DBOpen "SELECT     [Type Bayar].TypeFreight FROM         [Type Bayar] INNER JOIN                       [PO Order] ON [Type Bayar].TypeLoco = [PO Order].TypeLoco", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Tipe Pembayaran"
               Exit Function
            End If
       Case reDelMasterTransport:
            rcDelete.DBOpen "SELECT [PO Order].ID FROM         [PO Order] LEFT OUTER JOIN                       Transport ON [PO Order].ID = Transport.ID GROUP BY [PO Order].ID HAVING      ([PO Order].ID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            'MessageBox rcDelete.PrepareQuery
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Purchasing / Sales"
               Exit Function
            End If
            rcDelete.DBOpen "SELECT TransData.ID FROM         Transport INNER JOIN                       TransData ON Transport.ID = TransData.ID GROUP BY TransData.ID HAVING      (TransData.ID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Voucher"
               Exit Function
            End If
       Case reDelMasterGudangCust:
            rcDelete.DBOpen "SELECT     TransData.[GDG ID] FROM         TransData INNER JOIN                       [Gudang Customer] ON TransData.[GDG ID] = [Gudang Customer].[GDG ID] GROUP BY TransData.[GDG ID] HAVING      (TransData.[GDG ID] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Voucher"
               Exit Function
            End If
       Case reDelMasterBank:
            rcDelete.DBOpen "SELECT     [Bank Partner].Account FROM         [PO Order] LEFT OUTER JOIN  [Bank Partner] ON [PO Order].Account = [Bank Partner].Account GROUP BY [Bank Partner].Account HAVING      ([Bank Partner].Account = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Bank Partner"
               Exit Function
            End If
       Case reDelMasterPartner:
            rcDelete.DBOpen "SELECT     PartnerID FROM         [PO Order] GROUP BY PartnerID HAVING      (PartnerID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Purchasing / Sales"
               Exit Function
            End If
            rcDelete.DBOpen "SELECT     PartnerId FROM         TransData GROUP BY PartnerId HAVING      (PartnerId = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Voucher"
               Exit Function
            End If
       Case reDelMasterAktiva:
            rcDelete.DBOpen "SELECT     [No Aktiva] FROM         [DTR Aktiva Tetap] WHERE     ([No Aktiva] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Aktiva"
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT     [Doc Reff] FROM         [Detail Journal] WHERE     ([Doc Reff] = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Detail Journal"
               Exit Function
            End If
            
       Case reDelMasterResources:
            rcDelete.DBOpen "SELECT TypeID FROM [Resources Table] GROUP BY TypeID HAVING      (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Master Resources"
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT TypeID FROM [wcenter_resources] GROUP BY TypeID HAVING  (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Workcenter"
               Exit Function
            End If
            
            rcDelete.DBOpen "SELECT     TypeID FROM         [Manufacture Stage] GROUP BY TypeID HAVING      (TypeID = N'" & ParameterMaster & "')", CNN, lckLockReadOnly
            If rcDelete.Recordcount <> 0 Then
               CekCascadeDeleteTable = "Manufacture Routing"
               Exit Function
            End If
            
End Select
rcDelete.CloseDB
End Function
