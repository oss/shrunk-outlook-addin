
<center>
 <h1>Shrunk Outlook Add-in</h1>
</center>

https://github.com/oss/Shrunk-Outlook-Add-In/assets/7038712/7274efe7-d5ff-4762-9d17-b1ed8009e6ca

# Goal
To easily insert images (specifically tracking pixels) to an Outlook Email.

# Features

- Insert tracking pixels to your Outlook Email
- Support for (an infinite!) number of tracking pixels
- Automatically detect invisible tracking pixels as you draft your email
    - Undos, redos, re-ordering, deletions, etc. are all reflected onto the task pane
- Quickly show the locations of your tracking pixels within your email body

# Steps to Test Add-in
1. Run `npm install` inside root directory. Once that's done, run `npm start`
2. Navigate to https://outlook.office.com/mail/
3. Open any email
4. Under the subject, there will be a few icons. Click square with four circles in it (Apps)
5. Find the extension and click show Task pane

More help to side-load:
https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=web#modern-outlook-on-the-web-and-new-outlook-on-windows-preview
