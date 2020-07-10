//
//  main.swift
//  mailAttachment
//
//  Created by David M. Reed on 7/9/20.
//  Copyright © 2020 David Reed. All rights reserved.
//

import ArgumentParser
import Foundation
import AppleScriptObjC


/// create AppleScript command to email one file as an attachment
/// - Parameters:
///   - sender: email account to use to send (must be a valid account in Mail app)
///   - recipient: email address of the recipient
///   - fileURL: URL of file to send
/// - Returns: String containing AppleScript command that when executed will send the email
func createAppleScript(emailSender: String, emailRecipient: String, fileURL: URL) -> String {

    let script = """
    set p to "\(fileURL.path)"
    set theAttachment to POSIX file p

    tell application "Mail"
        set theNewMessage to make new outgoing message with properties {subject:"attached file", sender:"\(emailSender)", content:"see attached\n", visible:true}

        tell theNewMessage
            make new to recipient at end of to recipients with properties {address:"\(emailRecipient)"}
        end tell
        tell content of theNewMessage
            try
                make new attachment with properties {file name:theAttachment} at after the last word of the last paragraph
                set message_attachment to 0
            on error errmess -- oops
                log errmess -- log the error
                set message_attachment to 1
            end try
            log "message_attachment = " & message_attachment
        end tell
        delay 5
        tell theNewMessage
            send
        end tell
    end tell
    """

    return script
}

/// for each subdirectory in directory, attach one file and email to the email address specified by subdirectory name
/// - Parameters:
///   - emailSender: email account to use to send (must be a valid account in Mail app)
///   - directory: URL of directory containing all the subdirectories that are email address names
func sendEmails(emailSender: String, directory: String) {
    let fm = FileManager()
    let url = URL(fileURLWithPath: directory)
    if let contents = try? fm.contentsOfDirectory(at: url, includingPropertiesForKeys: [.isDirectoryKey], options: [.skipsSubdirectoryDescendants, .skipsHiddenFiles]) {
        for content in contents {
            if let v = try? content.resourceValues(forKeys: [.isDirectoryKey]) {
                if let isDir = v.isDirectory, isDir {
                    let directoryContents = try? fm.contentsOfDirectory(at: content, includingPropertiesForKeys: nil, options: [.skipsHiddenFiles, .skipsSubdirectoryDescendants, .skipsPackageDescendants])
                    if let firstURL = directoryContents?.first {
                        let emailRecipient = content.lastPathComponent
                        let script = createAppleScript(emailSender: emailSender, emailRecipient: emailRecipient, fileURL: firstURL)

                        // execute the AppleScript
                        var error: NSDictionary?
                        if let scriptObject = NSAppleScript(source: script) {
                            let output = scriptObject.executeAndReturnError(&error)
                            if error != nil {
                                print("error: \(error!)")
                            }
                            else if let msg = output.stringValue {
                                print(msg)
                            }
                        }
                    }
                }
            }
        }
    }
}

struct EmailSender: ParsableCommand {
    @Argument() var sender: String

    func run() {
        let fm = FileManager()
        let cwd = fm.currentDirectoryPath
        sendEmails(emailSender: sender, directory: cwd)
    }
}

EmailSender.main()
