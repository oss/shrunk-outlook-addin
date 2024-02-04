
<center>
 <h1>Shrunk Outlook Add-in</h1>
</center>

https://github.com/oss/Shrunk-Outlook-Add-In/assets/7038712/7274efe7-d5ff-4762-9d17-b1ed8009e6ca

# Goal
To easily insert images (specifically tracking pixels) to an Outlook Email.

# Features

- Insert tracking pixels to your Outlook Email
- Support for (an infinite?) number of tracking pixels
- Automatically detect invisible tracking pixels as you draft your email
    - Undos, redos, re-ordering, deletions, etc. are all reflected onto the task pane
- Quickly show the locations of your tracking pixels within your email body
- Prevents multiple inserts of the same tracking pixel (even when you click thrice super fast!)

# To Get Started
### Prerequisites
1. Node.js
2. NPM

### Steps
1. Clone the repo `git clone git@github.com:oss/Shrunk-Outlook-Add-In.git` and `cd` into the project.
2. Run `npm install`. Once installation is finished, run `npm start`.
3. Navigate to `https://outlook.office.com/mail/`
4. Open any email
5. Under the subject, there will be a few icons. Click the square with four circles inside of it (Apps)
6. Find the extension and click on it.
7. Follow the instructions.

Click [here](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?tabs=web#modern-outlook-on-the-web-and-new-outlook-on-windows-preview) for more information on side loading.

## Important Information

Side-loading only works for Windows machines.

`shrunk-dist-prod.zip` and `shrunk-dist-dev.zip` are distribution builds of this project that are built on every new tag creation (new release version. see `.github/workflows/ci.yml` for more details). They differ in their `manifest.xml` files: when grabbing assets, the base URL either becomes `go.rutgers.edu/outlook/assets` or `shrunk.rutgers.edu/outlook/assets`, where `shrunk.r.e` is the test instance of `go.rutgers.edu`. 
