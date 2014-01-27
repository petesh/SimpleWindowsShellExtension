SimpleWindowsShellExtension
===========================

This is a simple windows shell extension to determine the operation when more than 16 items are selected before and after the verb is invoked.

If you select <= 16 items, then initialize is invoked once, prior to the display of the menu item. If you select the menu item, it will perform the InvokeCommand

If you select > 16 items then initialize is invoked twice. Once prior to the display of the menu item, with 16 items. Then, if you select the menu item it will be invoked a second time with the full list of items, and the the InvokeCommand operation will be invoked.
