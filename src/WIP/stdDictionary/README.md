# `stdDictionary`

Dictionaries are one of the most important classes in VBA. 

We have numerous existing dictionary classes to draw inspiration off of:


[`clsTrickHash`](https://www.vbforums.com/showthread.php?788247-VB6-Hash-table) - Use of machine code IEnumVARIANT is pretty cool! Not Mac compatible.
[`VBA-Dictionary`](https://github.com/VBA-tools/VBA-Dictionary/blob/master/Dictionary.cls) - By tim hall, solid VBA implementation. Unlikely this is very efficient, but it is Mac compatible.
[`HashTable`](http://www.devx.com/vb2themax/Tip/19307)  - Faster because we store hashes. Not mac compatible yet. Also very memory hungry


## Best solution:

Reimplementation of HashTable to use a binary tree search system:

Key --> Hash --> Lookup in BinaryTree

BinaryTree implementation:

Type Node
  v as variant 'value
  key as variant
  index as long 'index in master
  h as long 'hash
  l as long 'index in master
  r as long 'index in master
End Type

master = [node1,node2,node3,node4]


            node1
           /     \
        node2   node3
        /   \   /   \
       0     0 0    node4

node1.l = node2.index = 2
node1.r = node3.index = 3
node2.l = 0
node2.r = 0
node3.l = 0
node3.r = node4.index = 4

## Conflict resolution?