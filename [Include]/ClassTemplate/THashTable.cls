VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The hashtable is one of the most powerful and mysterious data structures.
'Like God, it has many names in many different languages--dictionary,
'associative array, etc.  Unlike God, it actually gives you what you ask for.
'Hashtables are renowned for their O(1) search time;
'this means that the amount of time required to find something in a hashtable
'does not scale with the number of elements in it.
'This makes it faster to find items in than an array, tree, or
'similar data structure for large data sets.
'
'I intend to illustrate how a hashtable works by using an class in VBA code
'that I found a long time ago on the Internet.  This is a good example because
'it is written in a simple high-level language, adds an important data structure
'to a language that lacks it, and VBA is NEVER taught in Computer Science
'courses.  This article will teach you how a hashtable works, not do your
'homework for you.

'
'This implementation uses a system-level function that allows you to copy
'objects in memory as part of its "hash function".  Hash functions are described
'in more detail later in this tutorial.

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)


'Then, you need to set up default values for the hashtable size, maximum items
'that can occupy any slot in the hashtable (more on this later), and the amount
'to expand the hashtable by when it runs out of space.  While hashtables are
'time-efficient, they are not space-efficient.  They take up more than the
'minimum size needed to hold the data they store and expand in chunks.  That is
'why they are not suitable for developers of embedded applications with small
'amounts of data and limited amounts of memory.


' default values
Const DEFAULT_HASHSIZE = 1024
Const DEFAULT_LISTSIZE = 2048
Const DEFAULT_CHUNKSIZE = 1024


'Each item in the hashtable is stored in a special data type.  It needs to know
'the key that will be used to retreive the item, the value to store, and be
'able to link to another one of this data structure.  The last item is needed
'because the hashtable needs to link together items with the same "hashcode".
'There are an infinite number of possible keys but only a finite number of
'spaces within the hashtable.  A function is needed to map each key to a space
'in the hashtable that inevitably will map many of the infinite keys into a
'single finite space.  What happens if you attempt to insert more than one key
'that maps to the same slot in the hashtable? The hashtable needs some way to
'resolve these "collisions".  This hashtable strings along all of the items that
'map to the same slot in a list and therefore needs a pointer to the next item
'in the list.

Private Type SlotType
    key As String
    value As Variant
    nextItem As Long ' 0 if last item
End Type

'To work, this hashtable requires some private variables describing the
'parameters described above.  Note that I added a Collection (resizable array)
'to hold a duplicate copy of each of the keys.  When items are inserted into a
'hashtable, the order in which they appear may be different than the order in
'which they were entered.  This collection ensures that the keys can be returned
'to the user in the order they were entered.


' for each hash code this array holds the first element
' in slotTable() with the corresponding hash code
Dim hashTbl() As Long
' the array that holds the data
Dim slotTable() As SlotType

' pointer to first free slot
Dim FreeNdx As Long

' size of hash table
Dim m_HashSize As Long
' size of slot table
Dim m_ListSize As Long
' chunk size
Dim m_ChunkSize As Long
' items in the slot table
Dim m_Count As Long

' This keeps the keys in the order they were entered for calls to the Keys
' property
Dim m_Keys As Collection

'This hashtable offers developers the opportunity to make the keys
'case-insensitive and requires a variable to
'indicate whether the keys are case sensitive or
'not.


' member variable for IgnoreCase property
Private m_IgnoreCase As Boolean

' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty

Property Get IgnoreCase() As Boolean
    IgnoreCase = m_IgnoreCase
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
    If m_Count Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    m_IgnoreCase = newValue
End Property

'As in any good complex abstract data type, most of the work involved in
'maintaining the data structure is hidden from the user.  In the case of this hashtable, there is a
'series of private functions that do most of the
'heavy lifting.
'
'First, you need a function that creates or expands the hashtable.  Internally,
'the hashtable is represented as an array of
'SlotTypes, where, as mentioned earlier, each
'SlotType can form a linked list of other
'SlotTypes.  The ExpandSlotTable function takes the
'number of elements to add and creates or resizes
'the array.
 

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub

'
'Then, you need the function that maps keys to slots in the array.  This "hash
'function" should try to evenly distribute keys
'throughout the table so as to minimize the number
'of collisions.  This implementation looks at the
'ASCII codes of each character in the string that
'makes up the key and XORs them.  Inside another
'function, which actually performs the mapping,
'this number is then subjected to modular division
'by the size of the hashtable to ensure that it
'always maps to a valid slot.  While this function
'does not produce an ideal distribution of key-
'hashcode mappings, it is fairly simple to
'implement and has performed well for me in many
'different projects.
'

Private Function HashCode(key As String) As Long
    Dim lastEl As Long, i As Long
    
' copy ansi codes into an array of long
    lastEl = (Len(key) - 1) \ 4
    ReDim codes(lastEl) As Long
' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal key, Len(key)
    
' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        HashCode = HashCode Xor codes(i)
    Next
    
End Function

'As mentioned above, you need a function that converts the hashcode into
'something that refers to a location in the
'slotTable.  This function generates the hash and
'then looks at that location and finds one of three
'things: an empty slot, one item, or a list of
'items.  If it finds an empty slot, it either
'creates one or returns nothing depending on the
'option.  If it finds one or more items, it needs
'to go through the list of items with the same code
'until it finds the one it is looking for.  Here
'is the limitation of the hashtable.  While the
'hashtable can be an O(1) data structure, it can
'also be very inefficient--in the worst case, all
'items are stored in one big list and finding an
'item becomes an O(n) operation.

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal key As String, Optional Create As Boolean, Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    
' raise error if invalid key
    If Len(key) = 0 Then Err.Raise 1001, , "Invalid key"
    
' keep case-unsensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
' get the index in the hashTbl() array
    HCode = HashCode(key) Mod m_HashSize
' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
' exit if there is no item with that hash code
    Do While ndx
' compare key with actual value
        If slotTable(ndx).key = key Then Exit Do
' remember last pointer
        LastNdx = ndx
' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
    
' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, key, HCode, LastNdx
    Else
' signal that no item has been created
        Create = False
    End If
' this is the return value
    GetSlotIndex = ndx

End Function

'The previous function uses two new functions: GetFreeSlot and PrepareSlot.
'GetFreeSlot looks at the list of unused elements
'to see if any are left and gets a reference to the
'first one.  If none are available, it needs to
'add new ones.  Note that this implementation does
'not rehash all of the values in the hashtable and
'expand m_HashSize.  This means that inserting too
'many values in the hashtable will cause its
'performance to decline as lists become longer.
'In contrast, rehashing will scale better but can
'make certain element insertions very time
'consuming as the entire table will need to
'rehash.  We may add a modified version of this
'hashtable as a revision to illustrate the
'difference between the implementations and let
'you test the performance of each for your
'application.
'
'PrepareSlot simply sets the key of the selected SlotType to a specified value.

' return the first free slot

Private Function GetFreeSlot() As Long
' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
' use the first slot
    GetFreeSlot = FreeNdx
' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal key As String, ByVal HCode As Long, ByVal LastNdx As Long)
' assign the key
' keep case-sensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
    slotTable(Index).key = key
    
    If LastNdx Then
' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub

'Now, back to the public properties and methods.  You need a class constructor.

Private Sub Class_Initialize()
' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
    Set m_Keys = New Collection
End Sub

'This function uses SetSize to initialize the hashtable.  Again, a very
'straightforward function.

' initialize the hash table

Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long)
' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub

'You need a function to see if an element in the hashtable exists.  It uses the
'GetSlotIndex to see if the key appears in the
'hashtable.  To insert a key-value pair, you can
'use the GetSlotIndex function described above to
'set the key and then use the reference that is
'returned to set the value.  Note that you also
'need to update the collection of keys so that you
'have a copy of the keys sorted in the order they
'were inserted.

' check whether an item is in the hash table

Function Exists(key As String) As Boolean
    Exists = GetSlotIndex(key) <> 0
End Function

' add a new element to the hash table

Sub Add(key As String, value As Variant)
    Dim ndx As Long, Create As Boolean
    
' get the index to the slot where the value is
' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(key, Create)
    
    If Create Then
' the item was actually added
        If IsObject(value) Then
            Set slotTable(ndx).value = value
        Else
            slotTable(ndx).value = value
        End If
        
        m_Keys.Add key
    Else
' raise error "This key is already associated with an item of this collection"
        Err.Raise 457
    End If
End Sub

'You need functions to set or get a specific item in the hashtable.  These
'methods use the private functions described above
'to do all of the heavy lifting.

' the value associated to a key
' (empty if not found)

Property Get item(key As String) As Variant

    Dim ndx As Long
' get the index to the slot where the value is
    ndx = GetSlotIndex(key)
    If ndx = 0 Then
' return Empty if not found
    ElseIf IsObject(slotTable(ndx).value) Then
        Set item = slotTable(ndx).value
    Else
        item = slotTable(ndx).value
    End If
End Property

Property Let item(key As String, value As Variant)
    Dim ndx As Long
' get the index to the slot where the value is
' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
' store the value
    slotTable(ndx).value = value
End Property

Property Set item(key As String, value As Object)
    Dim ndx As Long
' get the index to the slot where the value is
' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
' store the value
    Set slotTable(ndx).value = value
End Property

'Then, you need functions that allow you to remove one or all of the elements of
'the hashtable.  The remove function needs to
'handle the same the same three cases as
'GetSlotIndex--a non-existant item, a single item,
'or an element of a list of items.  If it finds a
'single item, it only needs to put it in the list
'of free items.  If it finds an element of a list,
'it needs to fix the links of the item before it so
'that it bypasses the item being removed and then
'adds it to the list of free items.  Finally, it
'needs to maintain the collection of keys to
'maintain their sorted order.  RemoveAll does not
'take the time to remove each individual element,
'it just initializes a new hashtable and lets the
'garbage collection sweep away the old one.

' remove an item from the hash table

Sub Remove(key As String)
    Dim ndx As Long, HCode As Long, LastNdx As Long
    Dim i As Integer
    
    ndx = GetSlotIndex(key, False, HCode, LastNdx)
' raise error if no such element
    If ndx = 0 Then Err.Raise 5
    
    If LastNdx Then
' this isn't the first item in the slotTable() array
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
' this is the first item in the slotTable() array
' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).nextItem
    Else
' this is the only item in the slotTable() array
' for this hash code
        hashTbl(HCode) = 0
    End If
    
' put the element back in the free list
    slotTable(ndx).nextItem = FreeNdx
    FreeNdx = ndx
    
' Remove the item from the keys collection
    For i = m_Keys.Count To 1 Step -1
        If m_Keys.item(i) = key Then
            m_Keys.Remove (i)
        End If
    Next i
    
' we have deleted an item
    m_Count = m_Count - 1
    
End Sub

' remove all items from the hash table

Sub RemoveAll()
    SetSize m_HashSize, m_ListSize, m_ChunkSize
    
' Clear the keys collection
    Set m_Keys = New Collection
End Sub

'Count and keys are self-explanatory.
' the number of items in the hash table

Property Get Count() As Long
    Count = m_Count
End Property

' the array of all keys
' (VB5 users: convert return type to Variant)

Property Get Keys() As Variant
    Dim res() As Variant
    Dim i As Integer
    
    ReDim res(m_Keys.Count - 1)
    For i = 0 To m_Keys.Count - 1
        res(i) = m_Keys.item(i + 1)
    Next i
    
    Keys = res()
End Property


