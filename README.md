BaronSoft.com - Online RPG Engine v050 – 08/04/2003

License Agreement

This is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License or GNU Lesser General Public License as published by the Free Software Foundation version 2 and 2.1 of the License respectively. Please refer to the individual source files to see which software agreement it is released under.

This software is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License or GNU Lesser General Public License for more details.

You should have received a copy of the GNU General Public License and GNU Lesser General Public License along with this library; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307 USA

Notes about the License Agreement Changes
	
	ORE is now distributed under the GNU Public License. The full license agreements can be found, in separate files, in this release. Basically the GNU Public License means two major changes for how you can use ORE. 
1.	Any distribution of ORE must include all source code, derived or original. This means that any games released, using OREv5 or higher, must include all of the source code. The source code can be provided through any number of methods such as in the actual distribution of the game, or provided from a link to a web page. This is meant to foster sharing among the ORE community, making ORE a even better engine.
2.	The restrictions on commercial projects have been relaxed. Meaning you can charge for access to an ORE server. Also you can charge a distribution fee for an ORE client. You cannot charge for any actual software derived from ORE.  

One final note, The GNU agreements require that the author be listed in any distributions of ORE. Please also include a reference to www.baronsoft.com so that the original project may be found.

Graphics

	The graphics included with the engine are from various sources outside of the ORE project. When possible, permission was asked. If you use any of these graphics in your project, you run the risk of infringing on other people’s copyrights. Please, keep this in mind.

A special thanks to the Argentum staff, they donated some of their graphics to the engine.
	
Release Information

Release 0.5.0 – 08/04/2003
	Fist non-test release. This release still needs the Sound, Inventory, Combat, and Stat systems implemented.

Test Release 4.5 – 05/12/2003
	This release builds off of TR4 adding NPCs, and the start of an Item system. Also Tile Exits have been added (like OREv040) so you no longer HAVE to use the scripting engine for exits. The map editor does not yet allow NPC, Item, and Tile Exit placement, but hopefully Fredrik will update his map editor shortly. The TileEngineDemoX has been updated to show how to use the new features.

Test Release 4 – 05/12/2003
	This release is intended to test the new script engine as well as show off the new coding structure. The server has seen the biggest overhaul from the last release, adding several new class files and increasing the object-oriented level. NPC, Inventory, and Combat are still to come.
	This release also includes Fredrik Alexandersson’s map editor, which has been made an official part of ORE now.

Test Release 3 – 02/28/2003
	This release is intended to test the basic client-server system. Both the client and server are at a very basic state. Only chatting, and basic player code are included. NPCs, objects, and everything else will come once the foundation is tested.
	Please send feedback on the results of your testing.

Test Release 2.5 – 02/19/2003
	This is an updated test release of the TileEngineX. It is being released to help people currently working on map editors. It has the following changes.
1.	New file structure. Folders and paths have been switched around. ORE will use this new file structure from now on.
2.	Map_Save and Map_Load now take a number as a parameter. Maps will be saved in the standard mapX.map file format where X is the number you passed to the functions.
3.	Engine is not capable of displaying blocked tiles.
4.	The TileEngineXDemo has been expanded to respond to blocked tiles.
5.	Map_Bounds_Get has been added to get the size of the current map.
6.	Map_Edges_Blocked_Set has been added to easily block all the edges around the map so players cannot see past the map. 
7.	The Char system has been changed in several ways making and old code that manipulated characters non functional.

Test Release 2 – 01/04/2003
	This is a pre-release version of the OREv5. It only includes the graphics engine called TileEngineX. 
The hope of this release is two folds. First, the engine requires more testing across a wide range of hardware platforms. Second, being a one-man project I require help to get this engine going. I am requesting that the ORE community help with the project. Specifically a map editor needs to be programmed. I estimate that a good map editor would take me weeks to program. That time could be better spent on the client-server code.
Whatever person or group decides to take on this project and completes it will of course get full credit and copyrights. The approved map editor will be distributed as part of the ORE package.

Test Release 1 – 10/24/2002
This is a pre-release version of the OREv5. It only includes a semi-completed graphics engine called TileEngineX. The engine has been precompiled due to its uncompleted state.
The hope of this release is to test the engine across a wide range of hardware platforms.

Compiling Notes

This project requires Visual Basic 6 to compile and run. The following outside references are required to compile the engine.

All
-DirectX8 for Visual Basic Type Library (dx8vb.dll) for network/graphics/sound

Client Only
-ActiveMovie control type library (quartz.dll) for sound engine.

Server Only
-Microsoft Script Control 1.0 (msscript.ocx) for script engine.

Support Community

Questions? Comments? Please visit the forums at www.baronsoft.com. Note, you must register to see the ORE support forums.
