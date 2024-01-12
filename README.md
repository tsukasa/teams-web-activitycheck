# Teams Web Activity Check
A user script that adds [Idle Detection][idle_detection_api] support to [Microsoft Teams' web client][msteams].

## Description
If you are using [Microsoft Teams][msteams] in a web browser, you might have come across this issue: Once you leave the tab or focus another application, Teams stops seeing you as active and will eventually set your status to "Away" despite you being active and working in other applications.

To prevent this, this user script adds support for the [Idle Detection API][idle_detection_api] to the web client. When the user is active and the screen is unlocked, the user script will regularly send events to the web client so your status stays on "Available".

## Requirements
- A web browser that supports the [Idle Detection API][idle_detection_api].
- [Tampermonkey][tampermonkey] or a similar browser extension to load/inject the user script.

## Installation and Usage
Install the script via [Tampermonkey][tampermonkey], activate it and reload the Teams web client.

One time only, you need to click the Teams logo in the top bar. This will request permission to use the Idle Detector.

Once you grant the permission, you are good to go and the step does not need to be repeated unless you clear the site permissions for [Microsoft Teams][msteams].


[msteams]: https://teams.microsoft.com
[idle_detection_api]: https://developer.mozilla.org/en-US/docs/Web/API/Idle_Detection_API
[tampermonkey]: https://www.tampermonkey.net
