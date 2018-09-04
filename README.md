# Battleship
A fully OOP, Model-View-Presenter implementation, 100% VBA

## What is this?

Something to play with and have fun with, something to learn with, something to share, something to enhance and extend for fun, because we can, because VBA is fully capable of doing this, and because VBA devs can do open-source on GitHub, too!

## Do I need [Rubberduck](https://github.com/rubberduck-vba/Rubberduck) to use this code?

You don't. But you're definitely going to have a much better time with Rubberduck (although.. whether it's this project or another), be it only to enjoy navigating all these classes in a treeview with a customized folder hierarchy. You will not be able to run the unit tests without Rubberduck (`Assert` calls will fail to resolve), but you can absolutely run and explore this code without Rubberduck.

## How do I play?

You need a desktop install of Microsoft Excel with macros enabled. If macros are disabled, the title screen should look like this:

![system error: macros are disabled, abort mission, abort, abort, abort...](https://user-images.githubusercontent.com/5751684/45008183-2e057200-afcf-11e8-83b8-d3c0152b1070.png)

Otherwise, the first step is to pick a UI - at this point there's only a "Worksheet" UI, so you click it and you're taken to the "Game" screen, where you pick the grid you want to play in, knowing that **Player 1 always shoots first**:

!["new game" screen: pick a grid, pick AI opponent](https://user-images.githubusercontent.com/5751684/45008322-17abe600-afd0-11e8-8e3d-b8122fb2b586.png)

### AI Strategies

Just implementations of various strategies for winning a Battleship game. For now:

 - **Random**; shoots at random everywhere it can, until *all enemy ships are* found. Then, the heat is on. Ships may be adjacent.
 - **FairPlay**; shoots at random everywhere it can, until *an enemy ship is* found. Then proceeds to destroy that ship, then keeps shooting at random until it finds another ship to destroy, until it wins the game. Ships will not be adjacent.
 - **Merciless**; shoots *in random-ish patterns* targeting the center and/or the edges of the grid, until it finds a ship to sink. Then proceeds to destroy it, then resumes the hunt. Will not shoot a position where the smallest possible ship it's still looking for, wouldn't fit at that position horizontally or vertically. Tends to avoid shooting in positions adjacent to previous known hits if it's not hunting a ship down. Its ships will not be adjacent.

### Phase I: Ship Positioning

To play the worksheet UI (other implementations may work differently), you can follow the in-game instructions:

![Fleet deployment; action required: deploy aircraft carrier; click to preview, right-click to rotate, double-click to confirm](https://user-images.githubusercontent.com/5751684/45008702-209db700-afd2-11e8-9149-4caf597147a9.png)

To place a ship, select the location of its top-most, left-most position. Click anywhere in the grid to preview; if the preview isn't where you thought it would be, try rotating the ship by right-clicking. Double-click to confirm the position when you're ready to place the next ship - the ships you've placed will appear in the "Fleet Status" box:

![Fleet deployment; action required; deploy battlesihp; click to preview, ...](https://user-images.githubusercontent.com/5751684/45008774-75413200-afd2-11e8-8dc2-ebf8571da981.png)

Once you've placed all your ships, ...your AI opponent has already done the same and the game is ready for Player 1 to begin:

![Enemy fleet detected; all systems ready; double click in the enemy grid to fire a missile.](https://user-images.githubusercontent.com/5751684/45008878-103a0c00-afd3-11e8-84af-7f9692d0f67e.png)

### Phase II: Seek & Destroy

The goal is to find and sink all enemy ships before they find and sink all of yours.

If you're playing grid 2, you cross your fingers while the AI picks a position to begin the game; if you placed your ships in grid 1, you double-click a cell in grid 2, and then the AI will play.

![player 1 missed in D4, player 2 hit battleship (E5, horizontal) in F6](https://user-images.githubusercontent.com/5751684/45008999-b5ed7b00-afd3-11e8-8c24-72cbe238c608.png)

As the game progresses and you sink enemy ships, specifically *which* ships you've taken down will appear in the "acquired targets" box under the opponent's grid - 

![acquired battleship and submarine, merciless AI sunk cruiser and battleship, and is two hits shy of sinking my carrier](https://user-images.githubusercontent.com/5751684/45009072-1381c780-afd4-11e8-8f55-2cf38d965394.png)

Once a player has found and destroyed all 5 enemy ships, the game ends:

![game over - player 2 (merciless AI) wins, I never found its cruiser](https://user-images.githubusercontent.com/5751684/45009351-aff89980-afd5-11e8-9b0b-a6334de9dbeb.png)

The ships (and their respective length) are:

 - **Aircraft Carrier** (5)
 - **Battleship** (4)
 - **Submarine** (3)
 - **Cruiser** (3)
 - **Destroyer** (2)

### Sounds cool! Does it run on a Mac?

It wasn't written with MacOS in mind, so there are `kernel32.dll` invokes to `Sleep` and `CopyMemory`, among others. Might work, might not, might blow up, I've no idea. I wouldn't try it on a Mac.

## How do I contribute?

If you find a bug, or have a feature request, you will want to [open an issue](https://github.com/rubberduck-vba/Battleship/issues/new).

If you want to submit a [pull request](https://github.com/rubberduck-vba/Battleship/pulls) that closes an [open issue](https://github.com/rubberduck-vba/Battleship/issues), you'll need to fork the repository and work off a local clone of the files; open the `Battleship.xlsm` file in a desktop install of Microsoft Excel, load the VBE. Add new classes, new test modules and methods, new game modes, AI implementations, a new UI to play with, or enhancements to the `WorksheetView` - for best results, regularly export your files to the local git clone directory, *commit* the set of changes, *push* them to your fork, and make pull requests that focus on the feature it's for - if your pull request includes Rubberduck unit tests, it's even better!
