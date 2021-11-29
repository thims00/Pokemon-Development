Hello, everyone. I'm announcing the release of YAPE (Yet Another Pokemon Editor) version 0.9. This is, essentially, a pre-release of the coming version 1.0 (due out when it's out). I figured the community could use the new tool, and since all of the main features are complete and at least lightly tested, I'm releasing 0.9 now.


Main Features
-------------

YAPE works on the 3rd generation games (R/S/FR/LG/E) and edits the following:
- All of the base stats, EVs, etc. This includes a few I have not seen editable in any other editors (such as the level-up rate).
- Evolutions
- Usable TMs/HMs
- Learned attacks. (it even handles updating all the pointers automatically; adding/removing these for a pokemon is now extremely easy.)
- Pokedex entries (including the height, weight, size, and text)


Some possibly less-obvious features
-----------------------------------

Everything in the "Pokémon Selection" group can be used to select the active Pokémon. That means you can click on an entry in the evolution tree, enter one of the dex numbers, or select by name, etc.

You CAN change the number of learned attacks for a Pokémon. YAPE keeps all learned attacks for all Pokémon in the same space and only places a restriciton on the total across all Pokémon (to avoid overwriting important data by accident.) If you wish to add more learned attacks to say, bulbasaur, you would need to remove one from another Pokémon. You can add more moves to the "?" Pokémon this way. In future versions I plan to have the ability to specify a custom pointer for these for more advanced users, but beware that there will be no checking/automatic pointer updating on any such custom pointer values.
The text for the Pokédex entries is like the learned attacks. YAPE only tracks the total space for all text. If you want to make one entry longer, you would need to shorten another one.

It is possible to change what a Pokémon breeds to, but this must be done by editing evolutions. The game just backtracks through the evolution tree to determine what the result of breeding is. (There are a few special cases where it will stop before the beginning, such as breeding wobuffet without having it hold a lax insence, but everything else uses the beginning of the tree.) See the help file for more detailed info on how to do some more advanced things with this.


FAQ
---

- Is it possible to add/edit dex entries for the "?" Pokémon (the ones between Celebi and Treecko)?

- No. At least not in any useful manner. There is a limit of 386 dex entries. Technically, you could have an entry for them by changing the national dex numbers of these Pokémon to something between 1 and 386. However, this provides no real benefit over simply editing the stats/graphics of another existing Pokémon that already has a national dex number in the required range. See post #33 of this thread for all the gory details. If you want to make a hack with more than 386 Pokémon, it is strongly suggested that you wait for D/P tools and use those games as a base instead.


- I opened a ROM in YAPE, just changed one thing (or nothing at all), hit save, and lots of bytes changed in the ROM. What gives?

- The additional change you are seeing is probably normal. YAPE doesn't specifically track what was edited, it just saves all data every time you do a save. The things that are likely to change are:
the pokedex text/pointers (mainly in leaf green, as YAPE compacts all the entries and lg has some empty space in its dex text area)
The order of evolution data. For example, Eevee has 5 evolutions. It is possible that YAPE may change the order that these 5 evolutions appear in. (Evolution data is compacted/sorted internally in YAPE. When it writes the data back, it puts it in the sorted order rather than original order.)


Feature Requests
----------------

I'm listing the feature requests that I'm considering here. I have no current timelines for any of these features, so please do not ask when any of them will be done. In fact, I make no guarantees that any of these will be finished at all. They will only be added if I have the time and can do so in a sane, stable, and user-friendly fashion.

Considered for v2.0 or sooner. I have at least a general idea of how to handle all of these, but in many cases the implementation is non-trivial:
- Egg move editing. (Thanks to Teh Baro for info on this)
- Pokédex order editing.
- Simple method for expanding the space available for dex text and learned attacks. This will allow the addition of many more learned attacks and larger text entries without losing the safety and convenience of YAPE's automatic pointer updates.
- Various other polish and niceties in the UI.
- Importing/Exporting data as text.

Considered for post-v2 (Probably a long way off, if I ever get to it.):
- Graphics editing for Pokémon (sprite, pallete, position on battlefield, presence of shadow, etc.). General graphics editing should be done with a tile editor or unlz etc.
- Considered for separate companion tools to YAPE:
- Type strength/weakness editor.
- Attack editor (i.e. something that edits attack name, power, accuracy, etc.)
- Item editor

If you would like to see a feature that's NOT listed above, feel free to let me know. Be aware, though, that I've got a pretty large list already, so the odds of any new requests (aside from very minor ones) being added in the near future is small.


If you're having issues with YAPE
---------------------------------

If the program will not start due to some error, you probably have one of the following issues:
- You do not have the .Net framework 2.0 or higher installed. You can get this through windows update or from microsoft.com. It is free.
- You're missing some of the dlls required by YAPE. These are included in the distribution; make sure you keep them in the same folder as the main program.
- If anything in the UI is confusing, check the help files first. They contain much more detailed descriptions about what some of the options and stats mean.
- Currently this only edits the English versions of the games. If someone can track down the necessary offsets for other versions, I will add support for those as well.


Translators Wanted
------------------

I would love to be able to release 1.0 in a few more languages. If anyone is interested in translating it, let me know. I'm using ini files for all of the program's text, so you would not need to do any actual programming. Take a look in data\en-us\*.ini for an idea of what would be involved in the translation. (Translating the help file would be nice as well, but I realize that is more work...)


Constructive feedback is, of course, welcome.

Enjoy!
--Silver314