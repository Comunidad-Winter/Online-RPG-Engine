How to use the ORE Scripter
----------------------------

Loading
--------
Once the scripter is launched it will try to load the grh.dat file from the given graphics path.
If the file wasn´t found or the path doesn´t exist it will ask you to fill in the paths
(the sonfiguration window will be displayed) and then it will retry.

Adding parent nodes and grhs
-----------------------------
Once it finishes loading, you will have all grhs in the list to the left. Clicking on any item
will make the grh viewer display the graphic. On the right is the tree view. To add parent nodes
to the tree, just fill in the name and click the Create Node button. You can move around nodes
to nest them inside others by drag-dropping. To add a graphic to the tree, just drag-drop the
graphic index from the list on the left to the destination node. Grh indexes MUST be nested inside a parent node.

Setting default layers
-----------------------
The set default layer control will create a new tree elemt wich says the name of the layer to
be used. This default layer can be used either for a whole group (if it´s placed inside a group
node), or can be specyfic for one grh (by nesting the default layer node inside the grh index node).
This way, you can set a default layer for a whole group, and specyfic layers for certain grh inside
that same group which uses different layers.

Deleting tree elements
-----------------------
If you want to delete any alement from the tree just select that tree element and press Supr.
All elements nested inside will be removed too.

Saving / Loading files
-----------------------
File names are set by default (Grh Script.dat and Tile Script.dat). The path where these files are
located is the script path you entered in the configuration window.