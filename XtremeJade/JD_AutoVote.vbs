set vote = createobject("wscript.shell")

' Uncomment if you want to use specific browser / run the process in new window.

'vote.run "firefox.exe"

'wscript.sleep (3500)

vote.run "http://www.xtremejade.com/index.php"

wscript.sleep (9500)
vote.sendkeys chr(9)

wscript.sleep (3000)
vote.sendkeys chr(9)

wscript.sleep (3000)
vote.sendkeys ("your-name")

wscript.sleep (4500)
vote.sendkeys chr(9)

vote.sendkeys ("your-password")
wscript.sleep (3500)

vote.sendkeys chr(9)
wscript.sleep (3500)
vote.sendkeys"{Enter}"
wscript.sleep (9200)

call msgbox("Your Login was Successful!")

wscript.sleep (4500)
vote.run "http://www.xtremejade.com/menu/vote"

wscript.sleep (12500)
vote.run "http://www.xtremejade.com/menu/vote2"

call msgbox("Now, vote for the server!")

' Uncomment the lines, if the voting system is broken.

'wscript.sleep (3500)
'vote.run "http://www.xtremetop100.com/out.php?site=1132316066"

'call msgbox("Wait some time for the second vote!")

'wscript.sleep (37500)
'vote.run "http://www.gtop100.com/topsites/Jade-Dynasty/out/81815"

wscript.quit