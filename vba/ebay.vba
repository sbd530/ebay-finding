Sub GetItemsFromEbay()
    Dim strURL As String, responseTxt As String
    Dim xDoc As New MSXML2.DOMDocument60, itemCount As Integer, XPath As String, cntr As Integer
    Dim shOutput As Worksheet, totalItems, totalPages, pageCounter, UPCAndMotherCateg As String
    Dim dataArray()
    'SpeedOn
    pageCounter = 1
    Set shOutput = Sheets("Search Results")
   
    shOutput.Range("dataTable").ClearContents
    ' Creating the URL along with the filters
    strURL = "https://svcs.ebay.com/services/search/FindingService/v1?OPERATION-NAME=findCompletedItems" & _
        "&SERVICE-VERSION=1.0.0" & _
        "&SECURITY-APPNAME=" & Range("F1") & _
        "&RESPONSE-DATA-FORMAT=XML" & _
        "&REST-PAYLOAD" & _
        "&keywords=" & Range("Keyword") & _
        "&itemFilter(0).name=LocatedIn" & _
        "&itemFilter(0).value=" & Range("LocatedIn") & _
        "&itemFilter(1).name=ListingType" & _
        "&itemFilter(1).value=" & Range("ListingType") & _
        "&itemFilter(2).name=MinPrice" & _
        "&itemFilter(2).value=" & Range("MinPrice") & _
        "&itemFilter(2).paramName=Currency" & _
        "&itemFilter(2).paramValue=" & Range("Currency") & _
        "&itemFilter(3).name=MaxPrice" & _
        "&itemFilter(3).value=" & Range("MaxPrice") & _
        "&itemFilter(3).paramName=Currency" & _
        "&itemFilter(3).paramValue=" & Range("Currency") & _
        "&itemFilter(4).name=SoldItemsOnly" & _
        "&itemFilter(4).value=true" & _
        IIf(Range("Condition") = "NA", "", "&itemFilter(5).name=Condition&itemFilter(5).value=" & Range("Condition")) & _
        "&outputSelector(0)=SellerInfo" & _
        "&sortOrder=EndTimeSoonest"
   
    XPath = "/r:findCompletedItemsResponse/r:searchResult/r:item"
    xDoc.SetProperty "SelectionLanguage", "XPath"
    xDoc.SetProperty "SelectionNamespaces", "xmlns:r='http://www.ebay.com/marketplace/search/v1/services'"
   
    ' Loop to get all the details/UPC will be fetched separately
    Do
        Application.StatusBar = "Processing page " & pageCounter & "/" & totalPages
        strURL = strURL & "&paginationInput.pageNumber=" & pageCounter
   
        responseTxt = getValidItemDetails(strURL) ' The API call
   
        ' XML Processing
        xDoc.LoadXML responseTxt
        DoEvents
       
        If pageCounter = 1 Then
            ' Get total pages and total items in result for the first iteration only
            totalPages = CInt(xDoc.SelectSingleNode("/r:findCompletedItemsResponse/r:paginationOutput/r:totalPages").nodeTypedValue)
            totalItems = CInt(xDoc.SelectSingleNode("/r:findCompletedItemsResponse/r:paginationOutput/r:totalEntries").nodeTypedValue)
            ReDim dataArray(1 To totalItems, 1 To 10)
        End If
   
        itemCount = xDoc.SelectNodes(XPath).Length
   
        ' Loop over all the items to get the required details
        On Error Resume Next
        For cntr = 1 To itemCount
            Application.StatusBar = "Processing page " & pageCounter & "/" & totalPages & " Item:" & (pageCounter - 1) * 100 + cntr & "/" & totalItems
       
            dataArray((pageCounter - 1) * 100 + cntr, 1) = Range("Keyword")
            dataArray((pageCounter - 1) * 100 + cntr, 2) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:itemId").nodeTypedValue
            dataArray((pageCounter - 1) * 100 + cntr, 3) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:title").nodeTypedValue
            dataArray((pageCounter - 1) * 100 + cntr, 4) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:viewItemURL").nodeTypedValue
            dataArray((pageCounter - 1) * 100 + cntr, 5) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:sellerInfo/r:sellerUserName").nodeTypedValue
            '            dataArray((pageCounter - 1) * 100 + cntr, 5) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:productId").nodeTypedValue
            dataArray((pageCounter - 1) * 100 + cntr, 6) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:sellingStatus/r:convertedCurrentPrice").nodeTypedValue
            dataArray((pageCounter - 1) * 100 + cntr, 8) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:listingInfo/r:watchCount").nodeTypedValue
           
            UPCAndMotherCateg = GetItemUPCAndMotherCateg(dataArray((pageCounter - 1) * 100 + cntr, 2))
            dataArray((pageCounter - 1) * 100 + cntr, 7) = Split(UPCAndMotherCateg, "::")(0)
            dataArray((pageCounter - 1) * 100 + cntr, 9) = Split(UPCAndMotherCateg, "::")(1)
           
            dataArray((pageCounter - 1) * 100 + cntr, 10) = xDoc.SelectSingleNode(XPath & "[" & cntr & "]/r:listingInfo/r:endTime").nodeTypedValue
        Next cntr
        On Error GoTo 0
        pageCounter = pageCounter + 1
    Loop While pageCounter <= totalPages
   

   
    shOutput.Range("A2:J" & totalItems + 1) = dataArray
    shOutput.ListObjects("dataTable").Resize shOutput.Range("A1:J" & totalItems + 1)
    shOutput.Activate
       
    ' Adding the hyperlinks
    For cntr = 2 To totalItems + 1
        shOutput.Range("D" & cntr).Hyperlinks.Add shOutput.Range("D" & cntr), "" & shOutput.Range("D" & cntr).Text
    Next
   
    SpeedOff
End Sub
Function getValidItemDetails(strURL As String) As String
    Dim xmlReq As New MSXML2.XMLHTTP60, responseTxt As String

    ' Make a call to get the items
    With xmlReq
        .Open "GET", strURL, False
        .setRequestHeader "CONTENT-TYPE", "XML"
        .setRequestHeader "X-EBAY-SOA-GLOBAL-ID", "EBAY-US"
        .setRequestHeader "X-EBAY-SOA-OPERATION-NAME", "findCompletedItems"
        .setRequestHeader "X-EBAY-API-REQUEST-Encoding", "XML"
        .setRequestHeader "X-EBAY-SOA-SECURITY-APPNAME", Range("F1")
       
        xmlReq.send
       
        responseTxt = .responseText
    End With
    getValidItemDetails = responseTxt
End Function
'Option Explicit
Function GetItemUPCAndMotherCateg(ByVal strItemID As String) As String
    Dim xmlReq As New MSXML2.XMLHTTP60, strAPIUrl As String
    Dim responseTxt As String, postBody As String
    Dim xDoc As New MSXML2.DOMDocument60, getItemUPC As String, getItemMotherCateg As String
    Dim XPath As String
   
    postBody = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<GetItemRequest xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
        "<RequesterCredentials>" & _
        "<eBayAuthToken>" & Range("F4") & "</eBayAuthToken>" & _
        "</RequesterCredentials>" & _
        "<ErrorLanguage>en_US</ErrorLanguage>" & _
        "<WarningLevel>High</WarningLevel>" & _
        "<DetailLevel>ItemReturnAttributes</DetailLevel>" & _
        "<ItemID>" & strItemID & "</ItemID>" & _
        "</GetItemRequest>"

'"<OutputSelector>Item.ProductListingDetails.UPC</OutputSelector>" & _
        "<OutputSelector>Item.PrimaryCategory.CategoryName</OutputSelector>" & _

    'Set xmlReq = New WinHttp.WinHttpRequest
    strAPIUrl = "https://api.ebay.com/ws/api.dll"
    xmlReq.Open "POST", strAPIUrl, False
    xmlReq.setRequestHeader "X-EBAY-API-DEV-NAME", Range("F2")
    xmlReq.setRequestHeader "X-EBAY-API-CERT-NAME", Range("F3")
    xmlReq.setRequestHeader "X-EBAY-API-CALL-NAME", "GetItem"
    xmlReq.setRequestHeader "X-EBAY-API-SITEID", 0
    xmlReq.setRequestHeader "X-EBAY-API-REQUEST-Encoding", "XML"
    xmlReq.setRequestHeader "X-EBAY-API-COMPATIBILITY-LEVEL", "923"
    xmlReq.setRequestHeader "X-EBAY-API-APP-NAME", Range("F1")
   
    xmlReq.send (postBody)
   
    'Set objXML = New MSXML2.DOMDocument
    responseTxt = xmlReq.responseText
   
    XPath = "/r:GetItemResponse/r:Item/r:ProductListingDetails/r:UPC"
    xDoc.SetProperty "SelectionLanguage", "XPath"
    xDoc.SetProperty "SelectionNamespaces", "xmlns:r='urn:ebay:apis:eBLBaseComponents'"
    xDoc.LoadXML responseTxt
   
    On Error Resume Next
    getItemUPC = xDoc.SelectSingleNode(XPath).nodeTypedValue
    getItemMotherCateg = Split(xDoc.SelectSingleNode("/r:GetItemResponse/r:Item/r:PrimaryCategory/r:CategoryName").nodeTypedValue, ":")(0)
    GetItemUPCAndMotherCateg = getItemMotherCateg & "::" & getItemUPC
    On Error GoTo 0
    DoEvents
End Function
Sub SpeedOn()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
End Sub
Sub SpeedOff()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub