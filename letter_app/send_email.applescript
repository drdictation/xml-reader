
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {subject:"Andrew Branson - SVHM Endoscopy Referral"}
    try
        set theAccount to (first account whose email address is "chamara.basnayake@svha.org.au")
        set account of newMessage to theAccount
    end try
    make new recipient at newMessage with properties {email address:{address:"Endoscopy.Waitinglist@svhm.org.au"}}
    make new cc recipient at newMessage with properties {email address:{address:"office@focusgastro.com.au"}}

    make new attachment at newMessage with properties {file:POSIX file "/Users/cbasnayake/Documents/Microsaas/XML Reader/letter_app/generated_endoscopy/Andrew_Branson_SVHM_endoscopy_2026-10-13.pdf"}
    open newMessage
end tell
