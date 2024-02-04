# [How to use EVENT]

* 1.The concept is that dark red command(You can see the color distinction in the drop-down list of `MODE`) can generate events,and others receive events.Command can bound to event and executed according to whether that is established or not.
* 2.After the `Event` is established, it will always exist.You can use `RemoveEvent` to invalidate it
* 3.`Event` name can be named freely

> Example:

`MODE` | `EVENT` | `ACTION` | NOTE
------ | -------- | -------- | --------
`Match` | `Event1` | `s.png` | generate event
`Click` | `Event1` | `Left` | receive event
`RemoveEvent` | `Event1` |  | invalidate it
`RemoveEvent` | `Event1,Event2` |  | can multiple
`RemoveEvent` |  |  | empty means all

# [Mode Document]
> Explain what to input in the `ACTION` column

***

### ⭐`Click`
> Mouse click

* Action
  * @Action

> Action: use `Left`,`Right`,`Left_Down`,`Left_Up`,`Right_Down`,`Right_Up`

* Example
  * `Click` `Left_Down`

***

### ⭐`Calc(Calculator)`
> To do the math

* Action
  * @Formulas

> Formulas: like `x = 0`,`x = x + 1`,operators can use `+` `-` `*` `/`

* Example
  * `Calc` `x = 0`
  * `Calc` `y = 0`
  * `Calc` `x = x + y*2 + 1`
  * `Calc` `text = "DOC" + x`
  * `Calc` `y = y + random(0,10)`

> You can use these values outside of Calc by entering them like `{value}`.
> Example:

MODE | ACTION
------ | --------
`Move` | `{x},{y}`
`WriteClipboard` | `{text}`

***

### ⭐`Calc-Check`
> Check values.If it is true, the Event establish,otherwise is not

* Action
  * @Formulas

> Formulas: like `x > 1`.operators can use `==` `!=`(equality),`<` `>` `>=`(relational),`&&`(conditional AND),`||`(conditional OR)

* Example
  * `Calc-Check` `x == 0`
  * `Calc-Check` `y > 5 || x > 5`

***

### ⭐`Clear Screen`
> Clear marked on the screen

* Action
  * Empty

* Example
  * `Clear Screen`

***

### ⭐`Delay`
> Repeatedly execute

* Action
  * @Time

> Time: ms,1000 ms=1 second

* Example
  * `Delay` `1000`

***

### ⭐`Goto`
>Go to number of line

* Action
  * @Number
> Number: number of line

* Example
  * `Goto` `20`

***

### ⭐`Jump`
>Jump number of line

* Action
  * @Number
> Number: number of line

* Example
  * `Jump` `-3`

***

### ⭐`Key`
> Keyboard input

* Action
  * @Key
  * @Key,@Type
  * @Key,@Time

> Key: key in value , can use `A` ~ `Z`,`0` ~ `9`,`F1` ~ `F12`,`WIN` ...... \
> Type: `Down` or `Up` \
> Time: ms,1000 ms=1 second , the time keydown to keyup

* Example
  * `key` `A,Down`
  * `key` `B,500`
  * `key` `LEFT`

***

### ⭐`Loop`
> Repeatedly execute

* Action
  * @Times
  * Empty

> Times: run several times \
> Empty: always running

* Example
  * `Loop` `5`

***

### ⭐`Move`
> Move to screen point

* Action
  * @X,@Y
  * Empty

> X: screen point x value \
> Y: screen point y value \
> Empty: If set `Event`,it can use image matching point

* Example
  * `Move` `500,600`
  * `Move` `0,0`

***

### ⭐`ModifierKey`
> Modifier key

* Action
  * @Key

> Key: input like `modifier key|key,modifier` key use `CTRL`,`ALT`,`SHIFT`...

* Example
  * `ModifierKey` `CTRL|V`

***

### ⭐`Match,Match&Draw`
> Match and get image point \
> Match RGB: If your target needs to be distinguished by color,using this will have a better effect \
> Match&Draw: will be marked on the screen,just for test

* Action
  * @Path
  * @Path,@Threshold
  * @Path,@X,@Y,@Width,@Height
  * @Path,@X,@Y,@Width,@Height,@Threshold

> Path: image file path,use .png format \
> Threshold: `0.8`~`1.0`(def 0.9) \
> X: start point x value \
> Y: start point y value

* Example
  * `Match` `apple.png`
  * `Match` `apple.png,500,500,1420,580`

***

### ⭐`RemoveEvent`
> If you want invalidate `Event`, use this.

* Action
  * Empty
  * @Type

> Empty: direct invalidate it \
> Type: can use `PUSH`,then remove the first data for event.The event will be invalidated if all the data is removed

* Example
  * `RemoveEvent`
  * `RemoveEvent` `PUSH`

***

### ⭐`Offset`
> 

* Action
  * @X,@Y

> X: move x value \
> Y: move y value

* Example
  * `Offset` `+10,-20`
  * `Offset` `100,0`

***

### ⭐`PlaySound`
>Go to number of line

* Action
  * @Path
  * Empty

> Path: sound file path,use .wav format \
> Empty: default sound

* Example
  * `PlaySound` `sound.wav`

***

### ⭐`RandomTrigger`
> Random trigger Event

* Action
  * @Percentage

> Percentage: 1~100 %,1 = 1% chance trigger Event

* Example
  * `RandomTrigger` `33`

***

### ⭐`Run .exe`
>

* Action
  * @Path
> Path: .exe Path

* Example
  * `Run .exe` `C:\Users\Metro.exe`
  * `Run .exe` `C:\Program Files (x86)\......\chrome.exe -incognito`
  * `Run .exe` `CMD.exe /c C:\test.bat`

***

### ⭐`SendKeyDown`,`SendKeyUp`
> Keyboard input(for Game)

* Action
  * @Key

> Key: key value , can use `A` ~ `Z`,`0` ~ `9`,`F1` ~ `F12`...

* Example
  * `SendKeyDown` `RIGHT`
  * `SendKeyUp` `RIGHT`

***
        
### ⭐`WriteClipboard`
> Set text data to clipboard. \
> Can be used with CTRL+V for text input.

* Action
  * @Text

> Text: text data to clipboard

* Example
  * `WriteClipboard` `Hello!`

***
