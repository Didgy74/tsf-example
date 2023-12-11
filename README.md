# tsf-example
The purpose of this repository is to show to myself, and others,
how to use Win32 Text Services Framework to get 
integration with textual input on Windows. This includes integration with
the touch-screen keyboard, auto-correction, text suggestions and also support
for Windows Speech Recognition.

# Status
Not yet finished


# Appending text not initiated by a correction (SetText event)
When appending text like this, i.e when pressing a single character button, you need to update your
source text *and* your source selection values *before* sending the OnTextChange event to the TSF sink.
Below is an example of pushing a single character to the text-edit session.
In this case, two operations occur: TextChange to insert the char, and SelectionChange to move the cursor.

*Some uncertainty: I think this applies to all text events not initiated by TSF API, i.e outside SetText method.*
*Also I think this means that any OnTextChange event implies a possible change in selection.*

## Example that works
```cpp
DeferFn([=] {
	auto oldSelIndex = g_currentSelIndex;
	auto oldSelCount = g_currentSelCount;
	g_Replace(oldSelIndex, oldSelCount, 1, &character);
	g_textStore->text.Replace(oldSelIndex, oldSelCount, 1, &character);
	G_UpdateSelection(oldSelIndex + 1, 0);
	g_textStore->text.UpdateSelection(oldSelIndex + 1, 0);
	if (g_textStore->currentSink != nullptr) {
		auto& sink = *g_textStore->currentSink;
		auto hr = HResult_Helper::Ok;
		TS_TEXTCHANGE textChangeRange = {};
		textChangeRange.acpStart = oldSelCount;
		textChangeRange.acpOldEnd = oldSelIndex + oldSelCount;
		textChangeRange.acpNewEnd = oldSelIndex + 1;
		hr = (HResult_Helper)sink.OnTextChange(0, &textChangeRange);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
	}
});
```
## Example that does not work
Below is an example of something that **does not work**. Here the operation is
1. Modify source text
2. Send text event to TSF
3. Modify source selection
4. Send selection event to TSF
```cpp
DeferFn([=] {
	auto oldSelIndex = g_currentSelIndex;
	auto oldSelCount = g_currentSelCount;
	g_Replace(oldSelIndex, oldSelCount, 1, &character);
	g_textStore->text.Replace(oldSelIndex, oldSelCount, 1, &character);
	if (g_textStore->currentSink != nullptr) {
		auto& sink = *g_textStore->currentSink;
		auto hr = HResult_Helper::Ok;
		TS_TEXTCHANGE textChangeRange = {};
		textChangeRange.acpStart = oldSelCount;
		textChangeRange.acpOldEnd = oldSelIndex + oldSelCount;
		textChangeRange.acpNewEnd = oldSelIndex + 1;
		hr = (HResult_Helper)sink.OnTextChange(0, &textChangeRange);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
	}
	G_UpdateSelection(oldSelIndex + 1, 0);
	g_textStore->text.UpdateSelection(oldSelIndex + 1, 0);
	if (g_textStore->currentSink != nullptr) {
		auto& sink = *g_textStore->currentSink;
		auto hr = HResult_Helper::Ok;

		hr = (HResult_Helper)sink.OnSelectionChange();
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
	}
});
```

### But why
When inspecting the order in which TSF API calls method on our TextStore in response to our OnTextChange event, 
we can spot the following:
```
RequestLock begin
        GetSelection 1:0
        GetText 'h'
        OnEndEdit
RequestLock end
```
This happens when we only press the letter 'h' on the keyboard. Take note that TSF will call GetSelection first.
It is *critical* that this function returns the updated selection values. It *will* break if you try to instead update it 
with a OnSelectionChange event after the fact!

# SetText and SetSelection can happen multiple within one RequestLock segment
There doesn't seem to be a any real limit to how many editing-events can be triggered by the IME
during a RequestLock function. Here's a sequence of events that I captured:
```
RequestLock begin
        OnStartComposition
        SetText 7:0 'good day'
        GetText 'Have a good day'
        OnUpdateComposition
        OnEndComposition
        SetSelection 15:0
        OnStartComposition
        SetText 15:0 ' '
        GetText 'Have a good day '
        OnUpdateComposition
        OnEndComposition
        SetSelection 16:0
        GetActiveView - 1
        GetScreenExt - Visual data not up to date
        GetTextExt - Visual data not up to date
        GetActiveView - 1
        GetScreenExt - Visual data not up to date
        GetTextExt - Visual data not up to date
        GetSelection 16:0
        OnEndEdit
RequestLock end
```
Here you can notice that the editing happen in the order 
SetText -> SetSelection -> SetText -> SetSelection.

This might mean you need to capture a list of arbitrary chains of set-text + set-selection
operations that need to be deferred, depending on your architecture.

# Handling SetText in deferred context
Every call to SetText is immediately followed up by GetText. The same applies to SetSelection.
It may also be immediately followed up by GetTextExt.

## Doesn't work: Returing error in SetText and correcting later
The issue here is that SetSelection never calls if the preceeding SetText fails.
We gotta pretend that SetText works, so that it will call SetSelection after.

Any dispatch of the SetText+SetSelection events need to happen after SetSelection begins,
but perhaps happen before SetSelection ends? (SetText happens before SetSelection)

## Might work: Pretending SetText and SetSelection worked and then failing on GetScreenExt
