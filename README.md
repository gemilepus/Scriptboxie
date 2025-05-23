<p align="center">
<img width="96" align="center" src="Metro/package.ico" />
</p>

<h1 align="center">Scriptboxie</h1>
<p align="center">Scriptboxie allows you to effortlessly manage and create scripts. Can be used to streamline repetitive and time-consuming tasks.</p>

<br>


[![GitHub release](https://img.shields.io/github/release/gemilepus/Scriptboxie.svg)](https://github.com/gemilepus/Scriptboxie/releases) 
[![GitHub downloads](https://img.shields.io/github/downloads/gemilepus/Scriptboxie/total)](https://github.com/gemilepus/Scriptboxie/releases) 

## Features
- effortlessly manage and create scripts
- script recorder
- image search
## Prerequisite
- Windows 10 | 11

<h1 align="center">Download</h1>

Download available at <https://github.com/gemilepus/Scriptboxie/releases>.

If you like Scriptboxie, you can support it:

<a href='https://ko-fi.com/R6R8IQ1MD' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://storage.ko-fi.com/cdn/kofi2.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>
<br>
<a href='https://patreon.com/gemilepus' target='_blank'><img width='143' style='border:0px;width:143;' src='https://raw.githubusercontent.com/gemilepus/Scriptboxie/refs/heads/master/Metro/img/patreon.png' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>

<h1 align="center">How to use</h1>
<p align="center">Normally, I would record my actions first, and then make adjustments, such as setting how many times to repeat, </p>
<p align="center">adjusting the delay, etc. This way, I can quickly finish script. </p>
<br>

<p align="center">edit script: just enter the keyboard and mouse actions to be use.</p>
<p align="center">After completion, you can save it as a .txt file</p>
<p align="center">
 <img align="center" alt="Main" src="Doc/doc2.png" />
</p>

<p align="center">There are also image search, calculation... and other functions,</p>
<p align="center">which can be used with events to make some variety of actions.</p>
<p align="center">
 <img align="center" alt="Main" src="Doc/doc4.png" />
</p>

<p align="center">setting script: set hotkey for your scripts</p>
<p align="center">
 <img align="center" alt="Main" src="Doc/doc1.png" />
</p>

<p align="center">Notice</p>
<p align="center">When this display OFF,the hotkey will not work.It will change back to ON after clicking it. </p>
<p align="center">This is to ensure that the script is not executed by mistake.</p>
<p align="center">
 <img align="center" alt="Edit" src="Doc/doc3.png" />
</p>

<h1 align="center">Documentation</h1>

Available at <https://github.com/gemilepus/Scriptboxie/blob/master/Metro/Resources/Documentation.md>
or
<p align="center">
 <img align="center" alt="Edit" src="Doc/documentation.png" />
</p>

[f1]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s1.png
[f2]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s2.png
[f3]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s3.png
[f4]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s4.png
[f5]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s5.png
[f6]: https://github.com/gemilepus/Scriptboxie/blob/master/Doc/s6.png

<h1 align="center">Example</h1>
<p align="center">Automatically click the button</p>

| | |
| ------------- | ----------- |
| if you want to click the Submit button | [![][f1]][f1] |
| take a screenshot and use [![][f2]][f2] save it , like this picture| [![][f4]][f4] |
| then refer to this picture to set | [![][f5]][f5] Note:<br>1.The concept is that dark red command can generate events,and others receive events. <br>Command can bound to event and executed according to whether that  is established or not.<br>2.event name can be named freely |
| done | :) | 

<p align="center">Then let it alway running</p>

| | |
| ------------- | ----------- |
| refer to this picture to set :) &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;| [![][f6]][f6] Note:After the Event is established, it will always exist.You can use RemoveEvent to invalidate it |

<h1 align="center">Default functions & values</h1>

Name | Explanation | Example
---- | ---- | ---- 
`StartPosition_X` | Starting X-axis position of the mouse | Move `{StartPosition_X},{StartPosition_Y}`
`StartPosition_Y` | Starting Y-axis position of the mouse | 
`CurrentPosition_X` | Current X-axis position of the mouse | 
`CurrentPosition_Y` | Current Y-axis position of the mouse | 
`ClipboardText` | The current text in the clipboard | Calc `text =  ClipboardText + ","`
`random(a,b)` | Random Number between a and b | Calc `x = random(1,100)`

<h1 align="center">Screenshots</h1>
<p align="center">
 Script recorder
 <img align="center" alt="Edit" src="Doc/test2.gif" />
 Manage scripts
 <img align="center" alt="Edit" src="Doc/test.gif" />
</p>

<h1 align="center">WIKI</h1>

- [Example](https://github.com/gemilepus/Scriptboxie/wiki/Example)
  - [Click all the button on the screen](https://github.com/gemilepus/Scriptboxie/wiki/Example#click-all-the-button-on-the-screen)
  - [Mouse drag](https://github.com/gemilepus/Scriptboxie/wiki/Example#mouse-drag)
- [Note](https://github.com/gemilepus/Scriptboxie/wiki/Note)
  - [Make hotkeys only work in specific windows](https://github.com/gemilepus/Scriptboxie/wiki/Note)

## Credits
- MahApps.Metro - https://github.com/MahApps/MahApps.Metro
- feather - https://github.com/feathericons/feather
- opencvsharp - https://github.com/shimat/opencvsharp
- globalmousekeyhook - https://github.com/gmamaladze/globalmousekeyhook
