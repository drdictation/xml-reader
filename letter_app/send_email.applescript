
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {subject:"Patient Letter: Ms Matilda MAYNARD"}
    try
        set theAccount to (first account whose email address is "drchamarabasnayake@gmail.com")
        set account of newMessage to theAccount
    end try
    make new recipient at newMessage with properties {email address:{address:"admin@endomelb.com.au"}}
    make new cc recipient at newMessage with properties {email address:{address:"office@focusgastro.com.au"}}
    make new cc recipient at newMessage with properties {email address:{address:"clinic@jeanhailes.org.au"}}

    make new attachment at newMessage with properties {file:POSIX file "/Users/cbasnayake/Documents/Microsaas/XML Reader/letter_app/generated_letters/Ms_Matilda_MAYNARD_2026-04-17.pdf"}
    open newMessage
end tell
