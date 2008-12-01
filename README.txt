SHEETNODE
A module to host spreadsheets as Drupal nodes.

INSTALLATION
You need the SocialCalc JavaScript spreadsheet first.

To get the latest SocialCalc code, checkout from the SVN repo:
$ svn co -r 65 https://repo.socialtext.net:8999/svn/socialcalc/trunk socialcalc

Apply the patches inside the socialcalc folder:
$ cd socialcalc
$ patch -p0 < ../socialcalctableeditor-absolute.patch
$ patch -p0 < ../socialcalc-3-comment.patch

