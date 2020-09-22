
 DRAGON RIDER v0.25
 ------------------
 by Danny van der Ark 
 danny@slave-studios.co.uk

I release this version of Dragon Rider (Work in progress!) because I think it shows how important a good website like pscode.com is.
Although i've been programming for over 20 years, Most of the specific techniques and some routines originate directly from other users who made their contributions to pscode.com. Thanks to them me and other people can learn a great deal and saving me hours if not days researching some little detail or function.

Dragon Rider is just another project I started for fun, nothing too serious. It's not a full game yet, it only has 1 (test) level.
But I hope some people can learn from this, and hopefully get some feedback. I hope the code is not too complex. There are a lot of comments in the code, but for my own use, so excuse my spelling and or poor choice of words (english is not my mother tongue). 

If you want more explanation of certain parts, or if you've got some ideas, complaints, or whatever, send me an email at danny@slave-studios.co.uk (or raptom@hotmail.com - but i don't check that one often).


CONCEPT
-Not really. Complete non-sense: You're a dragon taking a nap, after a few hundred years you wake up because of this annoying sound of sheep barring (?!) just outside your lair. You decide to go out with no other goal than to burn them to the ground! ;) 

SOUND
-plays multiple soundeffects (and music) at the same time using DX7Sound
-Plays music in the background at the same time using media player (mp3 not included in this release to save some Mb's).

COLLISION DETECTION
-pixel collision detection between player and landscape
-pixel collision detection between bullets and enemies
-bounding box collision detection for enemies and landscape

GAMEPLAY
-player sprite rotates 360 degrees and is blitted using the FoxCBmp library.
-Player and bullets move using an angle and a speed factor.
-Special bonus when you hit 5 or 10 enemies with a single bullet.
-Fire doesn't do direct damage but sets the enemy on fire, causing it damage each round.
-Player can fire 4 different weapons. 1=small lightning bolt, 2=breathweapon, 3=big lightning bolt, 4=plasma bolt or something.
-Enemies can fire back too (some of them), but they currentlt shoot downwards, ie. they can't aim at the player yet.

ANIMATION
-enemies have different properties, besides the usual hitpoints, etc. Each enemy has it's own behaviour. For example: when an enemy has a weapon and spots the player he will aproach the player and shoot a bullet, then run away untill he can fire again. When the enemy has no more ammo he switches to the program 'fear' making him run for his life away from the player ;)
-bullets have different properties (amount of damage, sprites, some bullets are destroyed on impact, others lose energy as they hit more enemies, etc.)

CHARACTERS
-Sheep: you main objective. They have no special animation or weapons (duh). they just randomly graze about.
-Farmers: (the little brown guys) these are relatively harmless humans, who run away at the first sight of a great red dragon.
-Soldiers: (the blue guys), these run towards to player with the intention to attack (but I haven't given them weapons yet ;).
-Priests: (the ones in the red robes), these approach the player and fire some sort of 'plasma bolts' or whatever, they run away untill they can fire again. they have limited ammo and run like hell when they're out of it.

LEVEL
-The level graphics (384 x 2000 pixels) are a jpg, so it can be created in any program. This was done in photoshop (or any other package), but could have been done using Lightwave or Max for example to make it look really good.
-Simple built in Level editor (to 'paint' where obstructions are or which area's are passable, where what enemies are located, etc.)
-A low res version of the level (24x125 pixels) determines what are obstructions, what's passable and the location of monsters. It uses color codes, for example Red (rgb 255,0,0) is the player start position, green (rgb 0,255,0) are sheep, white (rgb 255,255,255) are rocks or non-passable obstructions, etc. this data is also used for enemy navigation & (easy and fast) collision detection.
- High resolution black & white image of the level (384x2000 pixels) is used for accurate pixel collision for the player sprite.

KEYBOARD
During gameplay, use Left & right cursor keys to turn around. Use Up & down to accelerate or de-accelerate.
Press 1,2,3 or 4 to fire a weapon. 

Press D - to show debug information (amount of 'tiles' or enemies/monsters still alive, amount of VFX & BUllets currently in game)
Press F - to show the FPS (frames per second), which also shows the lowest fps-count suffered, between brackets.
Press M - to kill the music (not included in this release)
Press Q - or escape to quit the game

Press E - for Level-Editor mode, during the edit mode you can select any of the things listed below and they 'paint' over the grid boxes using the left-mouse button.

	press 1 to clear a space in the level, ie. make it passable again.
	press 2 to create a 'rock' or make that part not-passable (by monster or player).
	
	press A to place a Soldier
	press B to place a Farmer
	press C to place a Sheep
	press D to place a Priest

	Press P to set the player start position
	Press S to save (overwrite the background_code.bmp)

	Press E again to return to the game (where you can scroll up/down to edit another part of the level).



enjoy...

Danny,--
