#define NOMINMAX
#define INITGUID

#include <Windows.h>
#include <Textstor.h>
#include <msctf.h>
#include <olectl.h>
#include <Richedit.h>
#include <InputScope.h>
#include <Ctffunc.h>

#include <iostream>
#include <string>
#include <cmath>
#include <concepts>
#include <vector>

int viewId = 1;
auto englishLangId = 0x0409;
int charSize = 28;
int widgetSize = 600;

struct InputScopeTest : public ITfInputScope {
	ULONG m_refCount = 1; // Start with one reference

	//
	// IUnknown start
	//
	HRESULT QueryInterface(REFIID riid, void** ppvObject) override {
		if (ppvObject == nullptr) {
			return E_INVALIDARG;
		}

		IUnknown* outPtr = nullptr;
		if (riid == IID_IUnknown || riid == IID_ITfInputScope) {
			outPtr = (IUnknown*)(ITfInputScope*)this;
		}

		if (outPtr != nullptr) {
			outPtr->AddRef();
			*ppvObject = outPtr;
			return S_OK;
		} else {
			*ppvObject = nullptr;
			return E_NOINTERFACE;
		}
	}

	ULONG AddRef() override {
		return InterlockedIncrement(&this->m_refCount);
	}

	ULONG Release() override {
		long val = InterlockedDecrement(&this->m_refCount);
		if (val == 0) {
			delete this;
		}
		return val;
	}
	//
	// IUnknown end
	//

	//
	// ITfInputScope start
	//
	HRESULT GetInputScopes(
		InputScope** out_inputScopes,
		UINT* out_count) override
	{
		if (out_count == nullptr || out_inputScopes == nullptr) {
			return E_INVALIDARG;
		}
		*out_inputScopes = (InputScope*)CoTaskMemAlloc(sizeof(InputScope) * 1);
		// Check if allocation failed.
		if (!*out_inputScopes) {
			*out_count = 0;
			return E_OUTOFMEMORY;
		}

		auto outInputScopeArray = *out_inputScopes;

		outInputScopeArray[0] = IS_TEXT;

		*out_count = 1;
		return S_OK;
	}
	HRESULT GetPhrase(BSTR** ppbstrPhrases, UINT* pcCount) override { return E_NOTIMPL; }
	HRESULT GetRegularExpression(BSTR* pbstrRegExp) override { return E_NOTIMPL; }
	HRESULT GetSRGS(BSTR* pbstrSRGS) override { return E_NOTIMPL; }
	HRESULT GetXML(BSTR* pbstrXML) override { return E_NOTIMPL; }
	//
	// ITfInputScope end
	//
};

struct TextStoreTest;

std::string g_internalString;
int g_currentSelIndex = 0;
int g_currentSelCount = 0;
void G_UpdateSelection(int selIndex, int selCount) {
	if (selIndex + selCount > g_internalString.size()) {
		std::abort();
	}
	g_currentSelIndex = selIndex;
	g_currentSelCount = selCount;
}
// Returns true if the string somehow changed.
template<class T>
bool g_Replace(int selIndex, int selCount, int newTextCount, T const* newText) {
	if (selIndex + selCount > g_internalString.size()) {
		std::abort();
	}

	auto oldStringSize = g_internalString.size();
	auto sizeDiff = newTextCount - selCount;

	// First check if we need to expand our storage.
	if (sizeDiff > 0) {
		// We need to move all content behind the old substring
		// To the right.
		g_internalString.resize(oldStringSize + sizeDiff);
		auto end = selIndex + selCount - 1;
		for (auto i = oldStringSize - 1; i > end; i -= 1)
			g_internalString[i + sizeDiff] = g_internalString[i];
	} else if (sizeDiff < 0) {
		// We need to move all content behind the old substring
		// To the left.
		auto begin = selIndex + selCount;
		for (auto i = begin; i < oldStringSize; i += 1)
			g_internalString[i + sizeDiff] = g_internalString[i];
		g_internalString.resize(oldStringSize + sizeDiff);
	}

	for (auto i = 0; i < newTextCount; i += 1)
		g_internalString[i + selIndex] = newText[i];

	return selCount != 0 || newTextCount != 0;
}
template<class T = char>
bool g_Replace(int newTextCount, T const* newText) {
	auto returnVal = g_Replace(g_currentSelIndex, g_currentSelCount, newTextCount, newText);
	G_UpdateSelection(g_currentSelIndex + newTextCount, 0);
	return returnVal;
}

TextStoreTest* g_textStore = nullptr;

enum class HResult_Helper : long long {
    Ok = S_OK,
    S_False = S_FALSE,
    E_InvalidArg = E_INVALIDARG,
    E_Pointer = E_POINTER,
    E_NoInterface = E_NOINTERFACE,
    REGDB_E_ClassNotReg = REGDB_E_CLASSNOTREG,
    Class_E_NoAggregation = CLASS_E_NOAGGREGATION,
    E_AccessDenied = E_ACCESSDENIED,
    E_Unexpected = E_UNEXPECTED,
};

enum class LockFlag {
    None = 0,
    Read = TS_LF_READ,
    ReadWrite = TS_LF_READWRITE,
};

struct TextStoreTest :
    public ITextStoreACP2,
    public ITfContextOwnerCompositionSink,
    public ITfTextEditSink
{
    ULONG m_refCount = 1; // Start with one reference

    HWND hwnd = nullptr;

    ITextStoreACPSink* currentSink = nullptr;
    DWORD currentSinkMask = 0;
    // The type of current lock.
    //   0: No lock.
    //   TS_LF_READ: read-only lock.
    //   TS_LF_READWRITE: read/write lock.
    LockFlag current_lock_type_ = LockFlag::None;
    DWORD pendingAsyncLock = 0;

	std::vector<TS_ATTRID> requestedAttrs;
	bool requestAttrWantFlag = false;

    [[nodiscard]] auto HasReadLock() const {
        return
            ((unsigned int)current_lock_type_ & TS_LF_READ) != 0 ||
            ((unsigned int)current_lock_type_ & TS_LF_READWRITE) != 0;
    }

    [[nodiscard]] auto HasWriteLock() const {
        return ((unsigned int)current_lock_type_ & TS_LF_READWRITE) != 0;
    }

    [[nodiscard]] auto HasAnyLock() const {
        return (unsigned int)current_lock_type_ != 0;
    }

    //
    // IUnknown start
    //
    HRESULT QueryInterface(REFIID riid, void** ppvObject) override {
        if (ppvObject == nullptr) {
            return E_INVALIDARG;
        }

        IUnknown* outPtr = nullptr;
        if (riid == IID_IUnknown || riid == IID_ITextStoreACP2) {
            outPtr = (IUnknown*)(ITextStoreACP2*)this;
        } else if (riid == IID_ITfContextOwnerCompositionSink) {
            outPtr = (ITfContextOwnerCompositionSink*)this;
        } else if (riid == IID_ITfTextEditSink) {
            outPtr = (ITfTextEditSink*)this;
        }

        if (outPtr != nullptr) {
            outPtr->AddRef();
            *ppvObject = outPtr;
            return S_OK;
        } else {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }
    }

    ULONG AddRef() override {
        return InterlockedIncrement(&this->m_refCount);
    }

    ULONG Release() override {
        long val = InterlockedDecrement(&this->m_refCount);
        if (val == 0) {
            std::abort();
        }
        return val;
    }
    //
    // IUnknown end
    //


    //
    // TextStoreTest start
    //
    HRESULT AdviseSink(REFIID iid, IUnknown* unknown, DWORD mask) override {
        if (iid != IID_ITextStoreACPSink)
            return E_INVALIDARG;



		if (FAILED(unknown->QueryInterface(IID_PPV_ARGS(&this->currentSink))))
			return E_UNEXPECTED;
        if (this->currentSink != nullptr) {
            if (this->currentSink == unknown) {
                this->currentSinkMask = mask;
                return S_OK;
            } else {
                return CONNECT_E_ADVISELIMIT;
            }
        }

        this->currentSinkMask = mask;
        return S_OK;
    }
    HRESULT UnadviseSink(IUnknown* punk) override {
        std::abort();
        return S_OK;
    }

    enum class RequestLock_LockFlags_Helper {
        Read = TS_LF_READ,
        ReadWrite = TS_LF_READWRITE,
        Sync = TS_LF_SYNC,
    };
    HRESULT RequestLock(DWORD newLockFlagsIn, HRESULT *out_hrSession) override {
		if (!this->currentSink)
			return E_FAIL;
		if (!out_hrSession)
			return E_INVALIDARG;
		if (this->HasReadLock()) {
			if ((newLockFlagsIn & TS_LF_SYNC) != 0) {
				// Can't lock synchronously.
				*out_hrSession = TS_E_SYNCHRONOUS;
				return S_OK;
			}
			// Queue the lock request.
			if (this->pendingAsyncLock != 0)
				std::abort();
			this->pendingAsyncLock = newLockFlagsIn;
			*out_hrSession = TS_S_ASYNC;
			return S_OK;
		}
		// Lock
		this->current_lock_type_ = (LockFlag)newLockFlagsIn;
		// Grant the lock.
		*out_hrSession = this->currentSink->OnLockGranted((DWORD)current_lock_type_);
		// Unlock
		this->current_lock_type_ = (LockFlag)0;

		// Handles the pending lock requests.
		if (this->pendingAsyncLock != 0) {
			std::cout << "Granting a pending lock" << std::endl;
			this->current_lock_type_ = (LockFlag)this->pendingAsyncLock;
			this->currentSink->OnLockGranted((DWORD)current_lock_type_);
		}
		this->pendingAsyncLock = 0;
		this->current_lock_type_ = (LockFlag)0;

		return S_OK;

    }
    HRESULT GetStatus(TS_STATUS* out_status) override {
        if (!out_status)
            return E_INVALIDARG;

		out_status->dwDynamicFlags = TS_SD_TKBPREDICTIONENABLE | TS_SD_TKBAUTOCORRECTENABLE;

        // We don't support hidden text.
        out_status->dwStaticFlags =
            TS_SS_NOHIDDENTEXT |
            TS_SS_TRANSITORY |
            TS_SS_TKBAUTOCORRECTENABLE |
            TS_SS_TKBPREDICTIONENABLE;
        return S_OK;
    }

    HRESULT QueryInsert(
        LONG acpTestStart,
        LONG acpTestEnd,
        ULONG cch,
        [[out]] LONG* pacpResultStart,
        /* out */ LONG* pacpResultEnd) override
    {
        std::abort();
        return E_NOTIMPL;
    }

    HRESULT GetSelection(
        ULONG selectionIndex,
        ULONG selectionBufferSize,
        TS_SELECTION_ACP* out_selections,
        ULONG* out_fetchedCount) override
    {
        if (!out_selections)
            return E_INVALIDARG;
        if (!out_fetchedCount)
            return E_INVALIDARG;
        if (!this->HasReadLock())
            return TS_E_NOLOCK;
		if (selectionIndex != TS_DEFAULT_SELECTION) {
			std::abort();
		}
        *out_fetchedCount = 0;
        if (selectionBufferSize > 0) {
            out_selections[0].acpStart = (LONG)g_currentSelIndex;
            out_selections[0].acpEnd = (LONG)g_currentSelIndex + (LONG)g_currentSelCount;
            out_selections[0].style.ase = TS_AE_END;
            out_selections[0].style.fInterimChar = FALSE;
            *out_fetchedCount = 1;
        }
        return S_OK;
    }

    HRESULT SetSelection(
		ULONG ulCount,
		TS_SELECTION_ACP const* pSelections) override
    {
        if (ulCount == 0)
			return S_OK;

		if (pSelections == nullptr) {
			return E_INVALIDARG;
		}

		auto const& selection = pSelections[0];
		if (selection.acpEnd > g_internalString.size()) {
			std::abort();
		}

    	G_UpdateSelection(selection.acpStart, selection.acpEnd - selection.acpStart);

        return S_OK;
    }

	HRESULT GetText(
		LONG acpStart,
		LONG acpEnd,
		WCHAR* out_chars,
		ULONG charsOutCapacity,
		ULONG* out_charsCount,
		TS_RUNINFO* out_runInfos,
		ULONG runInfosOutCapacity,
		ULONG* out_runInfosCount,
		LONG* out_nextAcp) override
    {
    	std::cout << "GetText || ";

    	if (!this->HasReadLock()) {
    		return TF_E_NOLOCK;
    	}
    	if (out_charsCount == nullptr || out_runInfosCount == nullptr || out_nextAcp == nullptr) {
    		return E_INVALIDARG;
    	}
    	if (acpStart < 0) {
    		return E_INVALIDARG;
    	}
    	if (acpEnd < -1) {
    		return E_INVALIDARG;
    	}
    	if (acpEnd >= (LONG)g_internalString.size()) {
    		std::abort();
    	}
    	*out_charsCount = 0;
    	*out_runInfosCount = 0;
    	*out_nextAcp = 0;

    	int internalTextSize = (int)g_internalString.size();
    	int startIndex = std::min(internalTextSize, (int)acpStart);
    	int endIndex = acpEnd;
    	if (endIndex == -1)
    		endIndex = internalTextSize;
    	endIndex = std::min(endIndex, internalTextSize);
    	auto charCount = endIndex - startIndex;
    	charCount = std::max(0, std::min(charCount, (int)charsOutCapacity));
    	endIndex = startIndex + charCount;

    	*out_nextAcp = endIndex;

    	// Copy over our characters
    	if (charsOutCapacity > 0 && out_chars != nullptr) {
    		for (int i = 0; i < charCount; i++) {
    			out_chars[i] = (WCHAR)g_internalString[i + startIndex];
    		}
    		// Set the variable that says how many characters we have output.
    		*out_charsCount = (ULONG)charCount;

    		std::cout.write(&g_internalString[startIndex], charCount);
    	}

    	// For now, all our text is assumed to be visible, so we output that.
    	if (runInfosOutCapacity > 0 && out_runInfos != nullptr) {
    		auto& out = out_runInfos[0];
    		out.type = TS_RT_PLAIN;
    		out.uCount = charCount;
    		*out_runInfosCount = 1;
    	}

    	std::cout << std::endl;
    	return S_OK;
    }

    HRESULT SetText(
        DWORD dwFlags,
        LONG acpStart,
        LONG acpEnd,
        WCHAR const* newText,
        ULONG newTextCount,
        TS_TEXTCHANGE* out_change) override
    {
        if (!this->HasWriteLock())
            return TF_E_NOLOCK;
        if (out_change == nullptr)
            return E_INVALIDARG;
        if (newText == nullptr && newTextCount != 0)
            return E_INVALIDARG;
        if (acpStart < 0)
            return TF_E_INVALIDPOS;
        if (acpEnd < 0)
            std::abort();

		g_Replace(acpStart, acpEnd - acpStart, (int)newTextCount, newText);

		out_change->acpStart = acpStart;
		out_change->acpOldEnd = acpEnd;
		out_change->acpNewEnd = acpStart + (LONG)newTextCount;

        return S_OK;
    }

    HRESULT GetFormattedText(
		LONG acpStart,
		LONG acpEnd,
		IDataObject** ppDataObject) override
    {
        std::abort();
        if (ppDataObject == nullptr)
            return E_INVALIDARG;
        return E_NOTIMPL;
    }

    HRESULT GetEmbedded(
        LONG acpPos,
        REFGUID rguidService,
        REFIID riid,
        IUnknown** ppunk) override
    {
        std::abort();
        return E_NOTIMPL;
    }

    /*
    HRESULT GetWnd(TsViewCookie vcView, HWND* out_Hwnd) override {
        if (out_Hwnd == nullptr) {
            return E_INVALIDARG;
        }
        *out_Hwnd = this->hwnd;
        return S_OK;
    }
     */

    HRESULT QueryInsertEmbedded(
        GUID const* pguidService,
        FORMATETC const* pFormatEtc,
        BOOL* pfInsertable) override
    {
        std::abort();
        return E_NOTIMPL;
    }

    HRESULT InsertEmbedded(
        DWORD dwFlags,
        LONG acpStart,
        LONG acpEnd,
        IDataObject* pDataObject,
        TS_TEXTCHANGE* pChange) override
    {
        std::abort();
        return E_NOTIMPL;
    }

    HRESULT InsertTextAtSelection(
        DWORD dwFlags,
        WCHAR const* pchText,
        ULONG cch,
        LONG* pacpStart,
        LONG* pacpEnd,
        TS_TEXTCHANGE* pChange) override
    {
        std::abort();
        return E_NOTIMPL;
    }


    HRESULT InsertEmbeddedAtSelection(
        DWORD dwFlags,
        IDataObject* pDataObject,
        LONG* pacpStart,
        LONG* pacpEnd,
        TS_TEXTCHANGE* pChange) override
    {
        std::abort();
        return E_NOTIMPL;
    }


    HRESULT RequestSupportedAttrs(
        DWORD dwFlags,
        ULONG attrBufferSize,
        TS_ATTRID const* attrBuffer) override
    {
		if (attrBufferSize != 0 && attrBuffer == nullptr) {
			return E_INVALIDARG;
		}

		this->requestAttrWantFlag = false;
        if ((dwFlags & TS_ATTR_FIND_WANT_VALUE) != 0) {
            this->requestAttrWantFlag = true;
        }

		this->requestedAttrs.clear();
		for (int i = 0; i < (int)attrBufferSize; i++) {
			auto const& attr = attrBuffer[i];
			if (attr == GUID_PROP_INPUTSCOPE) {
				this->requestedAttrs.push_back(attr);
			}
		}

        return S_OK;
    }

    HRESULT RetrieveRequestedAttrs(
        ULONG attrValsCapacity,
        TS_ATTRVAL* out_AttrVals,
        ULONG* out_AttrValsFetched) override
    {
		if (out_AttrValsFetched == nullptr) {
			return E_INVALIDARG;
		}
		if (attrValsCapacity == 0) {
			*out_AttrValsFetched = 0;
			return S_OK;
		}

		// If we reached this branch, it means capacity is not zero, in which case the pointer can't be null.
		if (out_AttrVals == nullptr) {
			return E_INVALIDARG;
		}

		*out_AttrValsFetched = (ULONG)this->requestedAttrs.size();
		if (*out_AttrValsFetched > attrValsCapacity) {
			*out_AttrValsFetched = attrValsCapacity;
		}

		for (int i = 0; i < *out_AttrValsFetched; i++) {
			auto& outAttr = out_AttrVals[i];
			auto const& requestedAttr = this->requestedAttrs[i];

			// This should be set to 0 when using TSF.
			outAttr.dwOverlapId = 0;
			outAttr.idAttr = requestedAttr;

			if (this->requestAttrWantFlag) {
				outAttr.varValue.vt = VT_UNKNOWN;
				outAttr.varValue.punkVal = new InputScopeTest;
			} else {
				outAttr.varValue.vt = VT_EMPTY;
			}
		}

        return S_OK;
    }

    HRESULT RequestAttrsAtPosition(
        LONG acpPos,
        ULONG attrBufferSize,
        TS_ATTRID const* attrBuffer,
        DWORD dwFlags) override
    {
		if (attrBufferSize != 0 && attrBuffer == nullptr) {
			return E_INVALIDARG;
		}

		this->requestAttrWantFlag = false;
		if ((dwFlags & TS_ATTR_FIND_WANT_VALUE) != 0) {
			this->requestAttrWantFlag = true;
		}

		this->requestedAttrs.clear();
		for (int i = 0; i < (int)attrBufferSize; i++) {
			auto const& attr = attrBuffer[i];
			if (attr == GUID_PROP_INPUTSCOPE) {
				this->requestedAttrs.push_back(attr);
			}
		}

		return S_OK;
    }

    HRESULT RequestAttrsTransitioningAtPosition(
        LONG acpPos,
        ULONG cFilterAttrs,
        TS_ATTRID const* paFilterAttrs,
        DWORD dwFlags) override
    {
        std::abort();
        return E_NOTIMPL;
    }

    HRESULT FindNextAttrTransition(
        LONG acpStart,
        LONG acpHalt,
        ULONG cFilterAttrs,
        TS_ATTRID const* paFilterAttrs,
        DWORD dwFlags,
        LONG* out_pAcpNext,
        BOOL* out_pFound,
        LONG* out_pFoundOffset) override
    {
        std::abort();

        if (!out_pAcpNext || !out_pFound || !out_pFoundOffset)
            return E_INVALIDARG;
        *out_pAcpNext = 0;
        *out_pFound = FALSE;
        *out_pFoundOffset = 0;
        return S_OK;
    }

    HRESULT GetEndACP(/* [out] */ LONG* pacp) override {
        std::abort();
        return E_NOTIMPL;
    }

    HRESULT GetActiveView(TsViewCookie* outPvcView) override {
        if (outPvcView == nullptr)
            return E_INVALIDARG;

        *outPvcView = viewId;
        return S_OK;
    }

    HRESULT GetACPFromPoint(
        TsViewCookie vcView,
        POINT const* ptScreen,
        DWORD dwFlags,
        LONG* out_Acp) override
    {
        if (out_Acp == nullptr) {
            return E_INVALIDARG;
        }

	    POINT point = {};
	    bool success = ClientToScreen(this->hwnd, &point);
	    if (!success) {
		    return E_FAIL;
	    }

		int x = ptScreen->x - point.x;
		int y = ptScreen->y - point.y;
		if (x != 0 || y != 0) {
			std::abort();
		}

        *out_Acp = (LONG)0;

        return S_OK;
    }

    HRESULT GetTextExt(
        TsViewCookie vcView,
        LONG acpStart,
        LONG acpEnd,
        RECT* out_rect,
        BOOL* out_fClipped) override
    {
        if (out_rect == nullptr || out_fClipped == nullptr) {
            return E_INVALIDARG;
        }

        POINT point = {};
        bool success = ClientToScreen(this->hwnd, &point);
        if (!success) {
            return E_FAIL;
        }

    	auto acpCount = acpEnd - acpStart;
    	if (acpCount < 0 || acpCount > g_internalString.size()) {
    		std::abort();
    	}

        out_rect->left = point.x + charSize * acpStart;
        out_rect->top = point.y;
        out_rect->right = out_rect->left + charSize * acpCount;
        out_rect->bottom = out_rect->top + charSize;


    	// Check if the total text bounds are outside our widget-size.
    	auto textWidth = charSize * g_internalString.size();
    	auto textHeight = charSize;
        *out_fClipped = false;
    	if (textWidth > widgetSize || textHeight > widgetSize) {
    		*out_fClipped = true;
    	}


        return S_OK;
    }

    HRESULT GetScreenExt(
        TsViewCookie vcView,
        RECT* out_rect) override
    {
		//return E_NOTIMPL;
        if (out_rect == nullptr) {
            return E_INVALIDARG;
        }

        POINT point = {};
        bool success = ClientToScreen(this->hwnd, &point);
        if (!success) {
            return E_FAIL;
        }

        out_rect->left = point.x;
        out_rect->top = point.y;
        out_rect->right = out_rect->left + widgetSize;
        out_rect->bottom = out_rect->top + widgetSize;

        return S_OK;
    }
    //
    // TextStoreTest end
    //

    //
    // ITfContextOwnerCompositionSink start
    //
    HRESULT OnStartComposition(ITfCompositionView* pComposition, BOOL* out_ok) override
    {
        if (pComposition == nullptr || out_ok == nullptr)
            return E_INVALIDARG;

        *out_ok = true;
        return S_OK;
    }

    HRESULT OnUpdateComposition(ITfCompositionView* pComposition, ITfRange *pRangeNew) override
    {
        return S_OK;
    }

    HRESULT OnEndComposition(ITfCompositionView* pComposition) override
    {
        return S_OK;
    }
    //
    // ITfContextOwnerCompositionSink end
    //


    //
    // ITfTextEditSink start
    //
    HRESULT OnEndEdit(
        ITfContext* context,
        TfEditCookie ecReadOnly,
        ITfEditRecord* pEditRecord) override
    {
        //std::cout << "OnEndEdit" << std::endl;
        return S_OK;
    }
    //
    // ITfTextEditSink end
    //
};

LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	bool textChanged = false;
	bool selChanged = false;
	auto oldSelIndex = g_currentSelIndex;
	auto oldSelCount = g_currentSelCount;
	TS_TEXTCHANGE textChangeRange = {};

    switch (msg) {
    	/*
    	case WM_GETOBJECT: {
    		if ((DWORD)lParam == UiaRootObjectId && perWindowDataBackend.accessProvider != nullptr) { // UI Automation request
    			return backendData.uiAutomationFnPtrs.UiaReturnRawElementProvider(
					hwnd,
					wParam,
					lParam,
					perWindowDataBackend.accessProvider);
    		}
    		break;
    	}
    	*/

		case WM_KEYDOWN: {
			if (wParam == VK_BACK) {
				if (g_currentSelCount != 0) {
					textChanged = g_Replace<char>(g_currentSelIndex, g_currentSelCount, 0, nullptr);
					selChanged = true;
					G_UpdateSelection(g_currentSelIndex, 0);
					textChangeRange.acpStart = oldSelIndex;
					textChangeRange.acpOldEnd = oldSelIndex + oldSelCount;
					textChangeRange.acpNewEnd = oldSelIndex;
				} else {
					// Selection count is 0...
					if (g_currentSelIndex != 0) {
						textChanged = g_Replace<char>(g_currentSelIndex - 1, 1, 0, nullptr);
						G_UpdateSelection(g_currentSelIndex - 1, 0);
						selChanged = true;
						textChangeRange.acpStart = oldSelIndex - 1;
						textChangeRange.acpOldEnd = oldSelIndex;
						textChangeRange.acpNewEnd = oldSelIndex - 1;
					}
					// If selection is zero and index is 0, then we do nothing.
				}
			}
			if (wParam == 'A') {
				if (GetKeyState(VK_CONTROL) & 0x8000) {
					// Select all text
					selChanged = true;
					G_UpdateSelection(0, (int)g_internalString.size());
				}
			}
		}

    	break;
        case WM_CHAR: {
			auto character = (char)wParam;
			if (character < 32) {
				// The first 32 values in ASCII are control characters.
				break;
			}
			textChanged = g_Replace(oldSelIndex, oldSelCount, 1, &character);
        	G_UpdateSelection(oldSelIndex + 1, 0);
			selChanged = true;
			textChangeRange.acpStart = oldSelIndex;
			textChangeRange.acpOldEnd = oldSelIndex + oldSelCount;
			textChangeRange.acpNewEnd = oldSelIndex + 1;

        }
    }

	if ((textChanged || selChanged) && g_textStore->currentSink) {
		auto& sink = *g_textStore->currentSink;
		auto hr = (HResult_Helper)sink.OnStartEditTransaction();
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}

		if (textChanged) {
			hr = (HResult_Helper)sink.OnTextChange(0, &textChangeRange);
			if (hr != HResult_Helper::Ok) {
				std::abort();
			}
		}
		if (selChanged) {
			hr = (HResult_Helper)sink.OnSelectionChange();
			if (hr != HResult_Helper::Ok) {
				std::abort();
			}
		}

		hr = (HResult_Helper)sink.OnEndEditTransaction();
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
	}

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {

	AllocConsole();
	freopen("CONOUT$", "w", stdout);

	// Register the window class
	WNDCLASSEX wc = {};
	wc.cbSize = sizeof(WNDCLASSEX);
	wc.lpfnWndProc = WindowProc;
	wc.hInstance = hInstance;
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpszClassName = "MyWindowClass";
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

	auto registerClassResult = RegisterClassEx(&wc);
	if (registerClassResult == 0) {
		std::cout << "Error when registering class" << std::endl;
	}

	// Initialize COM
	auto comInitResult = (HResult_Helper) CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (comInitResult != HResult_Helper::Ok) {
		std::abort();
	}

	// Create the window
	auto hwnd = CreateWindow(
		"MyWindowClass",
		"TSF Test window",
		WS_OVERLAPPEDWINDOW | WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, 1280, 800,
		nullptr,
		nullptr,
		hInstance,
		nullptr);
	ShowWindow(hwnd, nCmdShow);
	SendMessage(hwnd, EM_SETEDITSTYLE, SES_USECTF, SES_USECTF);

	auto hr = HResult_Helper::Ok;

	ITfInputProcessorProfiles *inputProcProfiles = nullptr;
	ITfInputProcessorProfileMgr *inputProcProfilesMgr = nullptr;
	{
		hr = (HResult_Helper) CoCreateInstance(
			CLSID_TF_InputProcessorProfiles,
			nullptr,
			CLSCTX_INPROC_SERVER,
			IID_PPV_ARGS(&inputProcProfiles));
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}

		void *tempPtr = nullptr;
		hr = (HResult_Helper) inputProcProfiles->QueryInterface(IID_ITfInputProcessorProfileMgr, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		inputProcProfilesMgr = (ITfInputProcessorProfileMgr *) tempPtr;
	}


	ITfThreadMgr *threadMgr = nullptr;
	ITfThreadMgr2 *threadMgr2 = nullptr;
	ITfSourceSingle *threadMgrSource = nullptr;
	{
		auto hr = (HResult_Helper) CoCreateInstance(
			CLSID_TF_ThreadMgr,
			nullptr,
			CLSCTX_INPROC_SERVER,
			IID_PPV_ARGS(&threadMgr));
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		void *tempPtr = nullptr;
		hr = (HResult_Helper) threadMgr->QueryInterface(IID_ITfThreadMgr2, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		threadMgr2 = (ITfThreadMgr2 *) tempPtr;

		hr = (HResult_Helper) threadMgr->QueryInterface(IID_ITfSourceSingle, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		threadMgrSource = (ITfSourceSingle*) tempPtr;
	}

	TfClientId clientId = {};
	hr = (HResult_Helper)threadMgr2->Activate(&clientId);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}


    ITfDocumentMgr* docMgr = nullptr;
    hr = (HResult_Helper)threadMgr2->CreateDocumentMgr(&docMgr);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }

    g_textStore = new TextStoreTest;
    g_textStore->hwnd = hwnd;

    ITfContext* outCtx = nullptr;
    TfEditCookie outTfEditCookie = {};
    hr = (HResult_Helper)docMgr->CreateContext(
        clientId,
        0,
        (ITextStoreACP2*)g_textStore,
        &outCtx,
        &outTfEditCookie);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }

    void* ctxSourceTemp = nullptr;
    hr = (HResult_Helper)outCtx->QueryInterface(IID_ITfSource, &ctxSourceTemp);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }
    auto ctxSource = (ITfSource*)ctxSourceTemp;

    DWORD adviseSinkIdentifier = {};
    hr = (HResult_Helper)ctxSource->AdviseSink(IID_ITfTextEditSink, (ITfTextEditSink*)g_textStore, &adviseSinkIdentifier);
    //hr = (HResult_Helper)ctxSource->AdviseSingleSink(clientId, IID_ITfTextEditSink, (ITfTextEditSink*)g_textStore);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }

    hr = (HResult_Helper)docMgr->Push(outCtx);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }

    if (g_textStore->currentSink != nullptr) {
        auto& sink = *g_textStore->currentSink;
        hr = (HResult_Helper)sink.OnStartEditTransaction();
        if (hr != HResult_Helper::Ok) {
            std::abort();
        }

        TS_TEXTCHANGE textChangeRange = {};
        textChangeRange.acpStart = 0;
        textChangeRange.acpOldEnd = 0;
        textChangeRange.acpNewEnd = 0;
        hr = (HResult_Helper)sink.OnTextChange(0, &textChangeRange);
        if (hr != HResult_Helper::Ok) {
            std::abort();
        }

        hr = (HResult_Helper)sink.OnSelectionChange();
        if (hr != HResult_Helper::Ok) {
            std::abort();
        }

        hr = (HResult_Helper)sink.OnLayoutChange(TS_LC_CHANGE, viewId);
        if (hr != HResult_Helper::Ok) {
            std::abort();
        }

        hr = (HResult_Helper)sink.OnEndEditTransaction();
        if (hr != HResult_Helper::Ok) {
            std::abort();
        }
    }

	/*
    hr = (HResult_Helper)threadMgr2->SetFocus(docMgr);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }
    */

    ITfDocumentMgr* oldDocMgr = nullptr;
	hr = (HResult_Helper)threadMgr->AssociateFocus(hwnd, docMgr, &oldDocMgr);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}

    MSG msg = {};
    while (true) {
        // We always wait for a new result
        int getResult = GetMessage(&msg, nullptr, 0, 0);
        if (getResult >= 0) {
            // We got a message
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        } else {
            // Error?
        }
    }

    return 0;

}