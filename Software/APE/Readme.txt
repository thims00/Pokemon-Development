 
			Advanced Palette Editor (aka APE) v1.3.3

		   Copyrigth © 2007 Andrea Sartori aka Mew2 aka HackMew

 ===========================================================================================
				        -DESCRIPTION-
 ===========================================================================================

			  A program to edit palettes in GBA ROMs.

 ===========================================================================================
				       -INSTRUCTIONS-
 ===========================================================================================

 1) Open your favourite ROM.

 2) It's time to load a palette. You have two choice: loading it from offset, or
    searching it.

 3) In the first case, just type the offset in the respective textbox. If the palette
    you're going to load is a LZ77 compressed one, check the compressed palette option.
    Click the Load button, then. Now, if everything went fine, you should see all the
    16 colors in the Actual Palette field, both in hex and color view. At this point, 
    we're going to really edit the palette. To make so, type for each color in the 
    Actual Palette the respective one in the Changed Palette. Once you're done, click the 
    Replace button to save the changes made.

 4) In the latter case, you have to type in the Actual Palette's textboxes all the 16
    colors you want to look for. Eventually, you can search for less than than 16 colors but
    you have to search at least for one color anyway. Once ready, click the Search button.
    If you find some results, the offset textbox will change to the offset where the
    searched palette actually is.  If you click the Search button again, the program will
    scan the rest of the ROM for other identical palettes. A "Find Next", after all. When 
    you finally find the right palette, follow the instructions on step 3 to edit it.

 ===========================================================================================
				       -TIPS'N'TRICKS-
 ===========================================================================================

 In this section we're going to learn some almost-hidden or not so explicit things:
 
 * The RGB/GBA Converter can be used to quickly convert a RGB color (e.g. #00FF00) to the
   respective GBA one (e.g. E003) and vice versa. You can access it from the Tools menu.
 
 * The APE's Color Picker is a Photoshop-like color picker with an integrated eyedropper
   and magnifier. It helps you choosing the right color, when you need. You can open it
   from the Tools menu or just pressing CTRL+T.

 * Actual/Changed palettes can be imported and/or exported via the Edit menu.
   Supported formats are: GPL (APE Palette, native), ACT (Adobe Color Table),
   PAL (PaintShop Palette) and TPL (Tile Layer Pro Palette). Nevertheless, you can also copy
   and clear them.

 * When copying/clearing a palette pay attention because you can choose two different
   options: All Palettes and Only Active Palette. I think that the name are self-explanitory.
   To change that option, rightclick on the Copy/Clear imagebutton: a small menu will popup.
   That's it! 

 * When typing in the Actual/Changed palette textboxes you will notice that when you reach
   the limit for each one (which is 4), if you try to write something else the cursor will
   go to the next textbox.

 * The Index near the Actual/Changed palette can be used to load 16 different palettes just
   changing the Index and following the procedure to load a palette. Keep in mind that 
   when Exporting/Importing a palette, it will we be always treated as a 256 color one. So,
   16 colors for each Index.

 * Doubleclicking on the little color preview under each colors in the Actual/Changed palette
   will cause the Color Picker to popup. If the Auto Copy option is enabled, the GBA color
   chosen will be automatically copied to the clipboard. Moreover, if both Auto copy and
   Auto Paste options are enabled, when closing the Color Picker, the last selected color
   will replace the color that you have chosen in the Actual/Changed palette.

 * The Bookmarks function is useful to store, as the name say, some bookmarks and load them
   when you need. Each game version will have his own bookmarks that you can add/delete/change
   with the user-friendly editor.
   NOTE: when giving a description to a bookmark, remember to put "<C>" in front if the palette
   is compressed. Example: "<C> Mewtwo Palette".

 * In the Bookmark Editor, when doubleclicking a bookmark from the list, the selected bookmark
   will be loaded.

 * Rightclicking on the bookmark list will cause the Bookmark Search to popup. To search
   something, just type what you want on the textbox: the first occurrence, if found, will be
   highlighted.

 * The Safe Mode is a function that checks silently, when you want to load/edit a palette, if
   the colors inside it are really colors or not. Colors, in fact, go from 0000 to FF7F: any
   higher value is not a color, obviously. The Safe Mode alerts you by changing the background
   color of the "wrong" colors. Anyway, when loading a palette make sure to check the 
   "Compressed Palette" checkbox if it's a compressed one. Otherwhise, you could get some
   "wrong" colors which shouldn't be supposed to be there.

 * If you type something in the textbox you will notice the autocomplete function.

 * The Gradient-o-matic is a gradient maker accessible from the Tools menu. Just like the name say,
   it helps you to create unlimited color gradients that fit with your desires.