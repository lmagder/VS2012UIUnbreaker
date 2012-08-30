VS2012UIUnbreaker
===================
VS2012UIUnbreaker is a Visual Studio 2012 extension that does two main things:

Removes the ALLCAPS from the menus. It does this by poking some stuff through reflection.

Restores the normal window chrome. This removes the custom window border and leaves the normal Aero Glass style one. It works by subclassing the window to hijack the non-client area messages and intercepting the code to change the window region. Also, it uses reflection to poke at the WPF controls in the window to remove the now redundant client-area title.

Before
-------
![Before](https://raw.github.com/lmagder/VS2012UIUnbreaker/master/before.png "Before")

After
-------
![After](https://raw.github.com/lmagder/VS2012UIUnbreaker/master/after.png "After")
