set vote = createobject("wscript.shell")

vote.run "http://www.xtremejade.com/index.php"

wscript.sleep (8500)
vote.sendkeys chr(9)

wscript.sleep (2000)
vote.sendkeys chr(9)

wscript.sleep (3000)
vote.sendkeys ("your-name")

wscript.sleep (3000)
vote.sendkeys chr(9)

vote.sendkeys ("your-password")
vote.sendkeys chr(9)
vote.sendkeys"{Enter}"
wscript.sleep (7200)

call msgbox("Your Login was Successful!")

wscript.sleep (3200)
vote.run "http://www.xtremejade.com/menu/vote"

wscript.sleep (7500)
vote.run "http://www.xtremejade.com/menu/vote2"

call msgbox("Now, vote for the server!")

' Uncomment the lines, if the voting system is broken.

'wscript.sleep (3500)
'vote.run "http://www.xtremetop100.com/out.php?site=1132316066"

'call msgbox("Wait some time for the second vote!")

'wscript.sleep (37500)
'vote.run "http://www.gtop100.com/topsites/Jade-Dynasty/out/81815"

wscript.quit