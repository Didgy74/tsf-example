#define NOMINMAX
#define INITGUID

#include <Windows.h>
#include <Textstor.h>
#include <msctf.h>
#include <olectl.h>
#include <Richedit.h>
#include <InputScope.h>
#include <Ctffunc.h>
#include <uiautomationcore.h>
#include <uiautomation.h>

#include <iostream>
#include <string>
#include <cmath>
#include <concepts>
#include <utility>
#include <vector>
#include <functional>
#include <optional>

namespace tsf_example {
	template<typename T>
	bool Replace(std::string& text, int selIndex, int selCount, int newTextCount, T const* newText) {
		if (selIndex > text.size()) {
			std::abort();
		}

		auto oldStringSize = text.size();
		auto sizeDiff = newTextCount - selCount;

		// First check if we need to expand our storage.
		if (sizeDiff > 0) {
			// We need to move all content behind the old substring
			// To the right.
			text.resize(oldStringSize + sizeDiff);
			auto end = selIndex + selCount - 1;
			for (auto i = oldStringSize - 1; i > end; i -= 1)
				text[i + sizeDiff] = text[i];
		} else if (sizeDiff < 0) {
			// We need to move all content behind the old substring
			// To the left.
			auto begin = selIndex + selCount;
			for (auto i = begin; i < oldStringSize; i += 1)
				text[i + sizeDiff] = text[i];
			text.resize(oldStringSize + sizeDiff);
		}

		for (auto i = 0; i < newTextCount; i += 1)
			text[i + selIndex] = (char) newText[i];

		return selCount != 0 || newTextCount != 0;
	}

	int outputIndentCount = 0;

	std::string OutputIndents() {
		std::string text;
		for (int i = 0; i < outputIndentCount; i++) {
			text.push_back('\t');
		}
		return text;
	}

	using UiaReturnRawElementProvider_FnT = LRESULT(WINAPI*)(HWND hwnd, WPARAM wParam, LPARAM lParam, IRawElementProviderSimple* el);
	using UiaHostProviderFromHwnd_FnT = HRESULT(WINAPI*)(HWND hwnd, IRawElementProviderSimple** ppProvider);
	struct UiAutomationFnPtrs {
		UiaReturnRawElementProvider_FnT UiaReturnRawElementProvider = nullptr;
		UiaHostProviderFromHwnd_FnT UiaHostProviderFromHwnd = nullptr;
	};
	UiAutomationFnPtrs uiAutomationFnPtrs = {};

	struct TextStoreTest;
	struct MyUiAccessProvider;
	MyUiAccessProvider* uiAccessProvider = nullptr;

	constexpr int viewId = 1;
	constexpr auto englishLangId = 0x0409;
	constexpr int charSize = 28;
	constexpr int widgetSize = 600;
}

template<class Runnable>
class Defer {
public:
	Defer() = default;

	Defer(Runnable const& in) noexcept: runnable{&in} {}

	Defer(Defer const&) = delete;

	Defer(Defer&&) = delete;

	Defer& operator=(Defer const&) = delete;

	Defer& operator=(Defer&&) = delete;

	void Execute() {
		if (!done) {
			(*runnable)();
			done = true;
		}
	}

	~Defer() {
		Execute();
	}

private:
	Runnable const* runnable = nullptr;
	bool done = false;
};

struct ReplaceString {
	std::string text;
	int selIndex = 0;
	int selCount = 0;

	template<typename T>
	auto Replace(int selIndex, int selCount, int newTextCount, T const* newText) {
		return tsf_example::Replace(this->text, selIndex, selCount, newTextCount, newText);
	}

	void UpdateSelection(int selIndex, int selCount) {
		if (selIndex + selCount > this->text.size()) {
			std::abort();
		}
		this->selIndex = selIndex;
		this->selCount = selCount;
	}
};

enum class CustomMessageEnum {
	CustomMessage = WM_APP + 1,
};

tsf_example::TextStoreTest* g_textStore = nullptr;

DWORD g_threadId = 0;
std::string g_internalString;
int g_currentSelIndex = 0;
int g_currentSelCount = 0;
std::vector<std::function<void()>> g_queuedFns;

void DeferFn(std::function<void()> item) {
	g_queuedFns.push_back(item);
	while (true) {
		// Posting can fail in some cases, like when the message-queue is full, so we just keep trying until it works.
		auto postResult = PostThreadMessageA(
			g_threadId,
			(UINT) CustomMessageEnum::CustomMessage,
			0,
			0);
		// The result is nonzero if it succeeds.
		if (postResult != 0)
			break;
	}
}


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
	if (selIndex > g_internalString.size()) {
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
		g_internalString[i + selIndex] = (char) newText[i];

	return selCount != 0 || newTextCount != 0;
}

template<class T = char>
bool g_Replace(int newTextCount, T const* newText) {
	auto returnVal = g_Replace(g_currentSelIndex, g_currentSelCount, newTextCount, newText);
	G_UpdateSelection(g_currentSelIndex + newTextCount, 0);
	return returnVal;
}

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
			outPtr = (IUnknown*) (ITfInputScope*) this;
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
		UINT* out_count) override {
		if (out_count == nullptr || out_inputScopes == nullptr) {
			return E_INVALIDARG;
		}
		*out_inputScopes = (InputScope*) CoTaskMemAlloc(sizeof(InputScope) * 1);
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

struct tsf_example::TextStoreTest :
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

	struct SetTextJob {
		int start;
		int count;
		std::string text;
	};
	struct SetSelectionJob {
		int start;
		int count;
	};
	enum class TextEditingJobType {
		SetText, SetSelection
	};
	struct TextEditingJob {
		TextEditingJobType type{};
		SetTextJob setText{};
		SetSelectionJob setSelection{};
	};
	std::vector<TextEditingJob> queuedTextEditingJobs_;

	void EnqueueTextEditingJob(SetTextJob job) {
		queuedTextEditingJobs_.push_back({TextEditingJobType::SetText, std::move(job), {}});
	};

	void EnqueueTextEditingJob(SetSelectionJob job) {
		queuedTextEditingJobs_.push_back({TextEditingJobType::SetSelection, {}, job});
	};

	bool visualDataUpToDate = true;

	ReplaceString text;

	[[nodiscard]] auto& GetString() const { return text.text; }

	auto& GetString() { return text.text; }

	[[nodiscard]] auto HasReadLock() const {
		return
			((unsigned int) current_lock_type_ & TS_LF_READ) != 0 ||
			((unsigned int) current_lock_type_ & TS_LF_READWRITE) != 0;
	}

	[[nodiscard]] auto HasWriteLock() const {
		return ((unsigned int) current_lock_type_ & TS_LF_READWRITE) != 0;
	}

	[[nodiscard]] auto HasAnyLock() const {
		return (unsigned int) current_lock_type_ != 0;
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
			outPtr = (IUnknown*) (ITextStoreACP2*) this;
		} else if (riid == IID_ITfContextOwnerCompositionSink) {
			outPtr = (ITfContextOwnerCompositionSink*) this;
		} else if (riid == IID_ITfTextEditSink) {
			outPtr = (ITfTextEditSink*) this;
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

	HRESULT RequestLock(DWORD newLockFlagsIn, HRESULT* out_hrSession) override {
		std::cout << tsf_example::OutputIndents() << "RequestLock begin";
		Defer deferredEndLine = [] { std::cout << std::endl; };

		if (!this->currentSink) {
			std::cout << " - Error";
			return E_FAIL;
		}

		if (!out_hrSession) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		if (this->HasReadLock()) {
			if ((newLockFlagsIn & TS_LF_SYNC) != 0) {
				// Can't lock synchronously.
				*out_hrSession = TS_E_SYNCHRONOUS;
				std::cout << " - Error: Tried to immediately lock but it's already locked";
				return S_OK;
			}
			// Queue the lock request.
			if (this->pendingAsyncLock != 0)
				std::abort();
			this->pendingAsyncLock = newLockFlagsIn;
			*out_hrSession = TS_S_ASYNC;
			std::cout << " - Queuing async lock";
			return S_OK;
		}
		// Lock
		this->current_lock_type_ = (LockFlag) newLockFlagsIn;
		deferredEndLine.Execute();
		tsf_example::outputIndentCount++;
		// Grant the lock.
		*out_hrSession = this->currentSink->OnLockGranted((DWORD) current_lock_type_);
		// Unlock
		this->current_lock_type_ = (LockFlag) 0;


		// Handles the pending lock requests.
		if (this->pendingAsyncLock != 0) {
			this->current_lock_type_ = (LockFlag) this->pendingAsyncLock;
			this->currentSink->OnLockGranted((DWORD) current_lock_type_);
		}
		this->pendingAsyncLock = 0;
		this->current_lock_type_ = (LockFlag) 0;
		tsf_example::outputIndentCount--;
		std::cout << tsf_example::OutputIndents() << "RequestLock end" << std::endl;


		// Run our deferred job stuff
		if (!this->queuedTextEditingJobs_.empty()) {
			auto jobsCopy = this->queuedTextEditingJobs_;
			this->queuedTextEditingJobs_.clear();
			Defer([=] {
				for (auto const& job: jobsCopy) {
					if (job.type == TextEditingJobType::SetText) {
						auto const& textJob = job.setText;
						g_Replace(textJob.start, textJob.count, (int) textJob.text.size(), textJob.text.data());
						std::cout << "Deferred: Current text = '" << g_internalString << "'" << std::endl;
					} else {
						auto const& selJob = job.setSelection;
						G_UpdateSelection(selJob.start, selJob.count);
						std::cout << "Deferred: Current selection = " << g_currentSelIndex << ":" << g_currentSelCount
								  << std::endl;
					}
				}

				g_textStore->visualDataUpToDate = true;
				if (g_textStore->currentSink != nullptr) {
					auto& sink = *g_textStore->currentSink;
					auto hr = (HResult_Helper) sink.OnLayoutChange(TS_LC_CHANGE, viewId);
					if (hr != HResult_Helper::Ok) {
						std::abort();
					}
				}
			});
		}

		return S_OK;
	}

	HRESULT GetStatus(TS_STATUS* out_status) override {
		std::cout << tsf_example::OutputIndents() << "GetStatus";
		Defer _ = [] { std::cout << std::endl; };

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
		LONG* pacpResultStart,
		LONG* pacpResultEnd) override
	{
		std::abort();
		return E_NOTIMPL;
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
		std::cout << tsf_example::OutputIndents() << "GetText";
		Defer _ = [] { std::cout << std::endl; };

		if (!this->HasReadLock()) {
			std::cout << " - Error";
			return TF_E_NOLOCK;
		}
		if (out_charsCount == nullptr || out_runInfosCount == nullptr || out_nextAcp == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (acpStart < 0) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (acpEnd < -1) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (acpEnd >= (LONG) text.text.size()) {
			std::abort();
		}
		*out_charsCount = 0;
		*out_runInfosCount = 0;
		*out_nextAcp = 0;

		auto& myString = this->GetString();

		int internalTextSize = (int) myString.size();
		int startIndex = std::min(internalTextSize, (int) acpStart);
		int endIndex = acpEnd;
		if (endIndex == -1)
			endIndex = internalTextSize;
		endIndex = std::min(endIndex, internalTextSize);
		auto charCount = endIndex - startIndex;
		charCount = std::max(0, std::min(charCount, (int) charsOutCapacity));
		endIndex = startIndex + charCount;

		*out_nextAcp = endIndex;

		// Copy over our characters
		if (charsOutCapacity > 0 && out_chars != nullptr) {
			for (int i = 0; i < charCount; i++) {
				out_chars[i] = (WCHAR) myString[i + startIndex];
			}
			// Set the variable that says how many characters we have output.
			*out_charsCount = (ULONG) charCount;

			std::cout << " '";
			std::cout.write(&myString[startIndex], charCount);
			std::cout << "'";
		}

		// For now, all our text is assumed to be visible, so we output that.
		if (runInfosOutCapacity > 0 && out_runInfos != nullptr) {
			auto& out = out_runInfos[0];
			out.type = TS_RT_PLAIN;
			out.uCount = charCount;
			*out_runInfosCount = 1;
		}

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
		std::cout << tsf_example::OutputIndents() << "SetText";
		Defer _ = [] { std::cout << std::endl; };
		if (!this->HasWriteLock()) {
			std::cout << " - Error";
			return TF_E_NOLOCK;
		}
		if (out_change == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (newText == nullptr && newTextCount != 0) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (acpStart < 0) {
			std::cout << " - Error";
			return TF_E_INVALIDPOS;
		}
		if (acpEnd < 0)
			std::abort();

		std::string newTextCopy;
		newTextCopy.resize(newTextCount);
		for (int i = 0; i < newTextCount; i++) {
			newTextCopy[i] = (char) newText[i];
		}

		// Print the new text.
		std::cout << " ";
		std::cout << acpStart;
		std::cout << ":";
		std::cout << acpEnd - acpStart;
		std::cout << " '";
		std::cout.write(newTextCopy.data(), newTextCopy.size());
		std::cout << "'";

		tsf_example::Replace(
			this->GetString(),
			acpStart,
			acpEnd - acpStart,
			(int) newTextCopy.size(),
			newTextCopy.data());

		this->EnqueueTextEditingJob(SetTextJob{
			.start = acpStart,
			.count = acpEnd - acpStart,
			.text = newTextCopy
		});
		this->visualDataUpToDate = false;

		out_change->acpStart = acpStart;
		out_change->acpOldEnd = acpEnd;
		out_change->acpNewEnd = acpStart + (LONG) newTextCount;

		return S_OK;
	}

	HRESULT GetSelection(
		ULONG selectionIndex,
		ULONG selectionBufferSize,
		TS_SELECTION_ACP* out_selections,
		ULONG* out_fetchedCount) override
	{
		std::cout << tsf_example::OutputIndents() << "GetSelection";
		Defer _ = [] { std::cout << std::endl; };
		if (out_selections == nullptr || out_fetchedCount == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}
		if (!this->HasReadLock()) {
			std::cout << " - Error";
			return TS_E_NOLOCK;
		}

		if (selectionIndex != TS_DEFAULT_SELECTION) {
			std::abort();
		}

		// Print our selection
		std::cout << " ";
		std::cout << this->text.selIndex;
		std::cout << ":";
		std::cout << this->text.selCount;

		*out_fetchedCount = 0;
		if (selectionBufferSize > 0) {
			out_selections[0].acpStart = (LONG) this->text.selIndex;
			out_selections[0].acpEnd = (LONG) this->text.selIndex + (LONG) this->text.selCount;
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
		std::cout << tsf_example::OutputIndents() << "SetSelection";
		Defer _ = [] { std::cout << std::endl; };

		if (ulCount == 0) {
			return S_OK;
		}

		if (pSelections == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		auto const selection = pSelections[0];
		if (selection.acpEnd > this->text.text.size()) {
			std::abort();
		}

		// Print the set selection
		std::cout << " ";
		std::cout << selection.acpStart;
		std::cout << ":";
		std::cout << selection.acpEnd - selection.acpStart;

		this->text.UpdateSelection(selection.acpStart, selection.acpEnd - selection.acpStart);
		this->EnqueueTextEditingJob(SetSelectionJob{
			.start = selection.acpStart,
			.count = selection.acpEnd - selection.acpStart});

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
		IUnknown** ppunk) override {
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
		for (int i = 0; i < (int) attrBufferSize; i++) {
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
		ULONG* out_AttrValsFetched) override {
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

		*out_AttrValsFetched = (ULONG) this->requestedAttrs.size();
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
		DWORD dwFlags) override {
		if (attrBufferSize != 0 && attrBuffer == nullptr) {
			return E_INVALIDARG;
		}

		this->requestAttrWantFlag = false;
		if ((dwFlags & TS_ATTR_FIND_WANT_VALUE) != 0) {
			this->requestAttrWantFlag = true;
		}

		this->requestedAttrs.clear();
		for (int i = 0; i < (int) attrBufferSize; i++) {
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

	HRESULT GetEndACP(LONG* pacp) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT GetActiveView(TsViewCookie* outPvcView) override {
		std::cout << tsf_example::OutputIndents() << "GetActiveView";
		Defer _ = [] { std::cout << std::endl; };

		if (outPvcView == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		*outPvcView = viewId;
		std::cout << " - ";
		std::cout << *outPvcView;
		return S_OK;
	}

	HRESULT GetACPFromPoint(
		TsViewCookie vcView,
		POINT const* ptScreen,
		DWORD dwFlags,
		LONG* out_Acp) override
	{
		std::cout << tsf_example::OutputIndents() << "GetACPFromPoint";
		Defer _ = [] { std::cout << std::endl; };
		if (out_Acp == nullptr || ptScreen == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		if (!this->visualDataUpToDate) {
			std::cout << " - Visual data not up to date";
			return E_FAIL;
		}

		POINT point = {};
		bool success = ClientToScreen(this->hwnd, &point);
		if (!success) {
			return E_FAIL;
		}

		// Window local coordinates.
		int x = ptScreen->x - point.x;
		int y = ptScreen->y - point.y;

		// Print out the point:
		std::cout << " - Point{" << x << " , " << y << "}";

		auto const charCount = (int) this->text.text.size();
		int closestIndex = 0;
		int minDistance = INT_MAX;
		for (int i = 0; i <= charCount; ++i) { // Note the <= here
			int charPosX = i * charSize;
			int distance = abs(x - charPosX);
			if (distance < minDistance) {
				minDistance = distance;
				closestIndex = i;
			}
		}

		std::cout << " - Index " << closestIndex;
		*out_Acp = (LONG) closestIndex;

		return S_OK;
	}

	HRESULT GetTextExt(
		TsViewCookie vcView,
		LONG acpStart,
		LONG acpEnd,
		RECT* out_rect,
		BOOL* out_fClipped) override
	{
		std::cout << tsf_example::OutputIndents() << "GetTextExt";
		Defer _ = [] { std::cout << std::endl; };
		if (out_rect == nullptr || out_fClipped == nullptr) {
			return E_INVALIDARG;
		}

		// Print which indices it is requesting
		std::cout << " ";
		std::cout << acpStart;
		std::cout << ":";
		std::cout << acpEnd - acpStart;

		if (!this->visualDataUpToDate) {
			std::cout << " - Visual data not up to date";
			return E_FAIL;
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
		std::cout << tsf_example::OutputIndents() << "GetScreenExt";
		Defer _ = [] { std::cout << std::endl; };
		if (out_rect == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		if (!this->visualDataUpToDate) {
			std::cout << " - Visual data not up to date";
			return E_FAIL;
		}

		POINT point = {};
		bool success = ClientToScreen(this->hwnd, &point);
		if (!success) {
			std::cout << " - Error";
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
	HRESULT OnStartComposition(ITfCompositionView* pComposition, BOOL* out_ok) override {
		std::cout << tsf_example::OutputIndents() << "OnStartComposition";
		Defer _ = [] { std::cout << std::endl; };
		if (pComposition == nullptr || out_ok == nullptr) {
			std::cout << " - Error";
			return E_INVALIDARG;
		}

		*out_ok = true;
		return S_OK;
	}

	HRESULT OnUpdateComposition(ITfCompositionView* pComposition, ITfRange* pRangeNew) override {
		std::cout << tsf_example::OutputIndents() << "OnUpdateComposition" << std::endl;
		return S_OK;
	}

	HRESULT OnEndComposition(ITfCompositionView* pComposition) override {
		std::cout << tsf_example::OutputIndents() << "OnEndComposition" << std::endl;
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
		ITfEditRecord* pEditRecord) override {
		std::cout << tsf_example::OutputIndents() << "OnEndEdit" << std::endl;
		/*
    	BOOL selectionChanged = false;
    	pEditRecord->GetSelectionStatus(&selectionChanged);

    	IEnumTfRanges* enumTfRanges = nullptr;
    	pEditRecord->GetTextAndPropertyUpdates(TF_GTP_INCL_TEXT, nullptr, 0, &enumTfRanges);
    	if (enumTfRanges != nullptr && selectionChanged) {
    		while (true) {
    			ULONG rangesFetched = 0;
    			ITfRange* ranges[5] = {};
    			enumTfRanges->Next(5, &ranges[0], &rangesFetched);
    			if (rangesFetched == 0)
    				break;
    			for (int i = 0; i < rangesFetched; i++) {
    				auto& range = *ranges[i];
    				WCHAR outText[128] = {};
    				ULONG textFetched = 0;
    				range.GetText(ecReadOnly, 0, &outText[0], 128, &textFetched);
    				int test = 0;
    			}
    		}

    	}
    	*/

		return S_OK;
	}
	//
	// ITfTextEditSink end
	//
};

struct MyUiAccessProvider_TextEdit :
	public IRawElementProviderSimple,
	public IRawElementProviderFragment,
	public ITextProvider,
	public IValueProvider
{
	ULONG m_refCount = 1; // Start with one reference

	int widgetId = 10;

	tsf_example::MyUiAccessProvider* parent_ = nullptr;
	explicit MyUiAccessProvider_TextEdit(tsf_example::MyUiAccessProvider* parent) {
		this->parent_ = parent;
	}

	//
	// IUnknown start
	//
	HRESULT QueryInterface(REFIID riid, void** ppvObject) override {
		if (ppvObject == nullptr) {
			return E_INVALIDARG;
		}

		IUnknown* outPtr = nullptr;
		if (riid == IID_IUnknown) {
			outPtr = (IUnknown*)static_cast<IRawElementProviderFragment*>(this);
		} else if (riid == IID_IRawElementProviderSimple) {
			outPtr = static_cast<IRawElementProviderSimple*>(this);
		} else if (riid == IID_IRawElementProviderFragment) {
			outPtr = static_cast<IRawElementProviderFragment*>(this);
		} else if (riid == IID_ITextProvider) {
			outPtr = static_cast<ITextProvider*>(this);
		} else if (riid == IID_IValueProvider) {
			outPtr = static_cast<IValueProvider*>(this);
		}

		if (outPtr != nullptr) {
			outPtr->AddRef();
			*ppvObject = outPtr;
			return S_OK;
		}

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	ULONG AddRef() override {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG Release() override {
		long val = InterlockedDecrement(&m_refCount);
		if (val == 0) {
			std::abort();
			delete this;
		}
		return val;
	}
	//
	// IUnknown end
	//

	//
	// IRawElementProviderSimple begin
	//
	HRESULT get_ProviderOptions(ProviderOptions* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = ProviderOptions_ServerSideProvider;
		return S_OK;
	}

	HRESULT GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;
		if (patternId == UIA_TextPatternId) {
			*pRetVal = static_cast<ITextProvider*>(this);
		} else if (patternId == UIA_ValuePatternId) {
			*pRetVal = static_cast<IValueProvider*>(this);
		}

		if (*pRetVal != nullptr) {
			(*pRetVal)->AddRef();
		}
		return S_OK;
	}

	HRESULT GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		pRetVal->vt = VT_EMPTY;
		if (propertyId == UIA_NamePropertyId) {
			pRetVal->vt = VT_BSTR;
			pRetVal->bstrVal = SysAllocString(L"Text edit test");
		} else if (propertyId == UIA_AutomationIdPropertyId) {
			pRetVal->vt = VT_BSTR;
			pRetVal->bstrVal = SysAllocString(L"Text edit test automation ID");
		} else if (propertyId == UIA_ControlTypePropertyId) {
			pRetVal->vt = VT_I4;
			pRetVal->intVal = UIA_TextControlTypeId;
		} else if (propertyId == UIA_IsTextPatternAvailablePropertyId) {
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_TRUE;
		} else if (propertyId == UIA_IsValuePatternAvailablePropertyId) {
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_TRUE;
		} else if (propertyId == UIA_IsKeyboardFocusablePropertyId) {
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_FALSE;
		} else if (propertyId == UIA_HasKeyboardFocusPropertyId) {
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_FALSE;
		} else if (propertyId == UIA_IsEnabledPropertyId) {
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_TRUE;
		} else if (propertyId == UIA_SizeOfSetPropertyId) {
			pRetVal->vt = VT_I4;
			pRetVal->intVal = 1;
		}

		return S_OK;
	}

	// Return null because this item are not directly hosted in a window.
	// Ref:
	// https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/Official%20Windows%20Platform%20Sample/UI%20Automation%20fragment%20provider%20sample/%5BC%2B%2B%5D-UI%20Automation%20fragment%20provider%20sample/C%2B%2B/ListItemProvider.cpp#L163
	HRESULT get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		(*pRetVal) = nullptr;
		return S_OK;
	}
	//
	// IRawElementProviderSimple end
	//

	//
	// IRawElementProviderFragment
	//
	HRESULT Navigate(
		NavigateDirection direction,
		IRawElementProviderFragment** pRetVal) override
	{
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		*pRetVal = nullptr;
		if (direction == NavigateDirection_Parent) {
			*pRetVal = (IRawElementProviderFragment*)this->parent_;
		}
		if (*pRetVal != nullptr) {
			(*pRetVal)->AddRef();
		}


		return S_OK;
	}

	// UI Automation gets this value from the host window provider, so supply NULL here.
	//	 Ref:
	//	 https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/Official%20Windows%20Platform%20Sample/UI%20Automation%20fragment%20provider%20sample/%5BC%2B%2B%5D-UI%20Automation%20fragment%20provider%20sample/C%2B%2B/ListProvider.cpp#L207
	HRESULT GetRuntimeId(SAFEARRAY** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		int rId[] = { UiaAppendRuntimeId, this->widgetId };

		SAFEARRAY* psa = SafeArrayCreateVector(VT_I4, 0, 2);
		if (psa == nullptr) {
			return E_OUTOFMEMORY;
		}
		for (LONG i = 0; i < 2; i++) {
			auto hr = (HResult_Helper)SafeArrayPutElement(psa, &i, &(rId[i]));
			if (hr != HResult_Helper::Ok) {
				return E_FAIL;
			}
		}
		*pRetVal = psa;
		return S_OK;
	}

	HRESULT get_BoundingRectangle(UiaRect* pRetVal) override;
	HRESULT GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override {
		// This could should likely always return null for our implementation.
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;
		return S_OK;
	}

	HRESULT SetFocus() override {
		return S_OK;
	}

	// We need to return ourselves since we are the root.
	// Ref:
	// https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/Official%20Windows%20Platform%20Sample/UI%20Automation%20fragment%20provider%20sample/%5BC%2B%2B%5D-UI%20Automation%20fragment%20provider%20sample/C%2B%2B/ListProvider.cpp#L260
	HRESULT get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = (IRawElementProviderFragmentRoot*)this->parent_;
		(*pRetVal)->AddRef();
		return S_OK;
	}
	//
	// IRawElementProviderFragment end
	//


	/*
	//
	// ITextEditProvider begin
	//
	HRESULT GetActiveComposition(ITextRangeProvider** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT GetConversionTarget(ITextRangeProvider** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}
	//
	// ITextEditProvider end
	//
	 */

	//
	// ITextProvider begin (Inherited by ITextEditProvider)
	//
	HRESULT GetSelection(SAFEARRAY** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT GetVisibleRanges(SAFEARRAY** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT RangeFromChild(IRawElementProviderSimple* childElement, ITextRangeProvider** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT RangeFromPoint(UiaPoint point, ITextRangeProvider** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT get_DocumentRange(ITextRangeProvider** pRetVal) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT get_SupportedTextSelection(SupportedTextSelection* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = SupportedTextSelection_Single;
		return S_OK;
	}
	//
	// ITextProvider end
	//


	//
	// IValueProvider begin
	//
	HRESULT SetValue(LPCWSTR val) override {
		std::abort();
		return E_NOTIMPL;
	}

	HRESULT get_Value(BSTR *pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = SysAllocString(L"Some value test");
		return S_OK;
	}

	HRESULT get_IsReadOnly(BOOL *pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = TRUE;
		return S_OK;
	}
	//
	// IValueProvider end
	//
};

struct MyUiAccessProvider_TextRange : public ITextRangeProvider {
	ULONG m_refCount = 1; // Start with one reference

	//
	// IUnknown start
	//
	HRESULT QueryInterface(REFIID riid, void** ppvObject) override {
		if (ppvObject == nullptr) {
			return E_INVALIDARG;
		}

		IUnknown* outPtr = nullptr;
		if (riid == IID_IUnknown) {
			outPtr = (IUnknown*)static_cast<ITextRangeProvider*>(this);
		} else if (riid == IID_ITextRangeProvider) {
			outPtr = static_cast<ITextRangeProvider*>(this);
		}
		if (outPtr != nullptr) {
			outPtr->AddRef();
			*ppvObject = outPtr;
			return S_OK;
		}

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	ULONG AddRef() override {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG Release() override {
		long val = InterlockedDecrement(&m_refCount);
		if (val == 0) {
			std::abort();
			delete this;
		}
		return val;
	}
	//
	// IUnknown end
	//
};

struct tsf_example::MyUiAccessProvider :
	public IRawElementProviderSimple,
	public IRawElementProviderFragment,
	public IRawElementProviderFragmentRoot
{
	ULONG m_refCount = 1; // Start with one reference

	MyUiAccessProvider_TextEdit* child = new MyUiAccessProvider_TextEdit(this);

	HWND hwnd = {};

	[[nodiscard]] auto Win32Hwnd() const { return this->hwnd; }

	//
	// IUnknown start
	//
	HRESULT QueryInterface(REFIID riid, void** ppvObject) override {
		if (ppvObject == nullptr) {
			return E_INVALIDARG;
		}

		IUnknown* outPtr = nullptr;
		if (riid == IID_IUnknown) {
			outPtr = (IUnknown*)static_cast<IRawElementProviderSimple*>(this);
		} else if (riid == IID_IRawElementProviderSimple) {
			outPtr = static_cast<IRawElementProviderSimple*>(this);
		} else if (riid == IID_IRawElementProviderFragment) {
			outPtr = static_cast<IRawElementProviderFragment*>(this);
		} else if (riid == IID_IRawElementProviderFragmentRoot) {
			outPtr = static_cast<IRawElementProviderFragmentRoot*>(this);
		}
		if (outPtr != nullptr) {
			outPtr->AddRef();
			*ppvObject = outPtr;
			return S_OK;
		}

		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	ULONG AddRef() override {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG Release() override {
		long val = InterlockedDecrement(&m_refCount);
		if (val == 0) {
			std::abort();
			delete this;
		}
		return val;
	}
	//
	// IUnknown end
	//

	//
	// IRawElementProviderSimple begin
	//
	HRESULT get_ProviderOptions(ProviderOptions* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = ProviderOptions_ServerSideProvider;
		return S_OK;
	}

	HRESULT GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;
		if (*pRetVal != nullptr) {
			(*pRetVal)->AddRef();
		}
		return S_OK;
	}

	HRESULT GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		pRetVal->vt = VT_EMPTY;
		if (propertyId == UIA_NamePropertyId) {
			pRetVal->vt = VT_BSTR;
			pRetVal->bstrVal = SysAllocString(L"Root Test");
			return S_OK;
		}

		return S_OK;
	}

	HRESULT get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) override {
		return tsf_example::uiAutomationFnPtrs.UiaHostProviderFromHwnd(Win32Hwnd(), pRetVal);
	}
	//
	// IRawElementProviderSimple end
	//


	HRESULT GetFocus(IRawElementProviderFragment** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		*pRetVal = this->child;
		(*pRetVal)->AddRef();
		return S_OK;
	}

	HRESULT ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		(*pRetVal) = nullptr;

		POINT point = { (LONG)x, (LONG)y };
		bool success = ScreenToClient(this->Win32Hwnd(), &point);
		if (!success) {
			return E_FAIL;
		}

		int widgetPosX = 0;
		int widgetPosY = 0;
		bool pointIsInside =
			point.x >= widgetPosX && point.x < tsf_example::widgetSize &&
			point.y >= widgetPosY && point.y < tsf_example::widgetSize;
		if (pointIsInside) {
			(*pRetVal) = static_cast<IRawElementProviderFragment*>(this->child);
		}

		if (*pRetVal != nullptr) {
			(*pRetVal)->AddRef();
		}

		return S_OK;
	}


	//
	// IRawElementProviderFragment
	//
	HRESULT Navigate(
		NavigateDirection direction,
		IRawElementProviderFragment** pRetVal) override
	{
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;

		if (direction == NavigateDirection_FirstChild || direction == NavigateDirection_LastChild) {
			*pRetVal = static_cast<IRawElementProviderFragment*>(child);
		}

		if (*pRetVal != nullptr) {
			(*pRetVal)->AddRef();
		}

		return S_OK;
	}

	// UI Automation gets this value from the host window provider, so supply NULL here.
	//	 Ref:
	//	 https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/Official%20Windows%20Platform%20Sample/UI%20Automation%20fragment%20provider%20sample/%5BC%2B%2B%5D-UI%20Automation%20fragment%20provider%20sample/C%2B%2B/ListProvider.cpp#L207
	HRESULT GetRuntimeId(SAFEARRAY** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;
		return S_OK;
	}

	// Because this is the root, this can return an empty rect, in which case the
	// UIA framework will use the size of the window.
	HRESULT get_BoundingRectangle(UiaRect* pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}

		*pRetVal = {};
		return S_OK;
	}

	HRESULT GetEmbeddedFragmentRoots(SAFEARRAY** pRetVal) override {
		// This could should likely always return null for our implementation.
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = nullptr;
		return S_OK;
	}

	HRESULT SetFocus() override {
		return S_OK;
	}

	// We need to return ourselves since we are the root.
	// Ref:
	// https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/Official%20Windows%20Platform%20Sample/UI%20Automation%20fragment%20provider%20sample/%5BC%2B%2B%5D-UI%20Automation%20fragment%20provider%20sample/C%2B%2B/ListProvider.cpp#L260
	HRESULT get_FragmentRoot(IRawElementProviderFragmentRoot** pRetVal) override {
		if (pRetVal == nullptr) {
			return E_INVALIDARG;
		}
		*pRetVal = static_cast<IRawElementProviderFragmentRoot*>(this);
		(*pRetVal)->AddRef();
		return S_OK;
	}
	//
	// IRawElementProviderFragment end
	//
};

HRESULT MyUiAccessProvider_TextEdit::get_BoundingRectangle(UiaRect* pRetVal) {
	if (pRetVal == nullptr) {
		return E_INVALIDARG;
	}

	POINT clientPos = {};
	bool success = ScreenToClient(this->parent_->Win32Hwnd(), &clientPos);
	if (!success) {
		std::abort();
	}

	pRetVal->left = clientPos.x;
	pRetVal->top = clientPos.y;
	pRetVal->width = tsf_example::widgetSize;
	pRetVal->height = tsf_example::widgetSize;

	return S_OK;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_GETOBJECT: {
			if ((DWORD) lParam == UiaRootObjectId) { // UI Automation request
				std::cout << "Windows requested UiAutomation" << std::endl;
				return tsf_example::uiAutomationFnPtrs.UiaReturnRawElementProvider(
					hwnd,
					wParam,
					lParam,
					tsf_example::uiAccessProvider);
			}
			break;
		}

		case WM_IME_STARTCOMPOSITION: {
			std::cout << "WM_IME_STARTCOMPOSITION" << std::endl;
			break;
		}

		case WM_KEYDOWN: {
			if (wParam == 'A') {
				if (GetKeyState(VK_CONTROL) & 0x8000) {
					// Select all text
					DeferFn([] {
						G_UpdateSelection(0, (int) g_internalString.size());
						g_textStore->text.UpdateSelection(0, (int) g_internalString.size());
						if (g_textStore->currentSink != nullptr) {
							auto hr = (HResult_Helper) g_textStore->currentSink->OnSelectionChange();
							if (hr != HResult_Helper::Ok) {
								std::abort();
							}
						}
					});
				}
			}
		}

			break;
		case WM_CHAR: {
			auto oldSelIndex = g_currentSelIndex;
			auto oldSelCount = g_currentSelCount;
			auto character = (char) wParam;
			if (character == '\b') {
				if (oldSelCount != 0) {
					DeferFn([=] {
						g_Replace(oldSelIndex, oldSelCount, 0, (char*) nullptr);
						g_textStore->text.Replace(oldSelIndex, oldSelCount, 0, (char*) nullptr);
						G_UpdateSelection(oldSelIndex, 0);
						g_textStore->text.UpdateSelection(oldSelIndex, 0);
						if (g_textStore->currentSink != nullptr) {
							TS_TEXTCHANGE textChangeRange = {};
							textChangeRange.acpStart = oldSelIndex;
							textChangeRange.acpOldEnd = oldSelIndex + oldSelCount;
							textChangeRange.acpNewEnd = oldSelIndex;
							auto hr = (HResult_Helper) g_textStore->currentSink->OnTextChange(0, &textChangeRange);
							if (hr != HResult_Helper::Ok) {
								std::abort();
							}
						}
					});
				} else if (oldSelIndex != 0) {
					// Count is 0
					DeferFn([=] {
						g_Replace(oldSelIndex - 1, 1, 0, (char*) nullptr);
						g_textStore->text.Replace(oldSelIndex - 1, 1, 0, (char*) nullptr);
						G_UpdateSelection(oldSelIndex - 1, 0);
						g_textStore->text.UpdateSelection(oldSelIndex - 1, 0);
						if (g_textStore->currentSink != nullptr) {
							TS_TEXTCHANGE textChangeRange = {};
							textChangeRange.acpStart = oldSelIndex - 1;
							textChangeRange.acpOldEnd = oldSelIndex + 1;
							textChangeRange.acpNewEnd = oldSelIndex - 1;
							auto hr = (HResult_Helper) g_textStore->currentSink->OnTextChange(0, &textChangeRange);
							if (hr != HResult_Helper::Ok) {
								std::abort();
							}
						}
					});
				}
				break;
			}
			if (character < 32) {
				// The first 32 values in ASCII are control characters.
				break;
			}

			DeferFn([=] {
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
					hr = (HResult_Helper) sink.OnTextChange(0, &textChangeRange);
					if (hr != HResult_Helper::Ok) {
						std::abort();
					}
				}
			});
		}
	}
	return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {

	g_threadId = GetCurrentThreadId();

	AllocConsole();
	freopen("CONOUT$", "w", stdout);

	bool setDpiAwareResult = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

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

	// Set your desired client area size
	int clientWidth = 1000;
	int clientHeight = 1000;

	RECT rect = { 100, 100, clientWidth, clientHeight };
	AdjustWindowRect(&rect, WS_OVERLAPPEDWINDOW, FALSE);

	// Create the window
	auto hwnd = CreateWindow(
		"MyWindowClass",
		"TSF Test window",
		WS_OVERLAPPEDWINDOW | WS_VISIBLE,
		rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top,
		nullptr,
		nullptr,
		hInstance,
		nullptr);
	ShowWindow(hwnd, nCmdShow);
	//SendMessage(hwnd, EM_SETEDITSTYLE, SES_USECTF, SES_USECTF);

	auto hr = HResult_Helper::Ok;

	ITfInputProcessorProfiles* inputProcProfiles = nullptr;
	ITfInputProcessorProfileMgr* inputProcProfilesMgr = nullptr;
	{
		hr = (HResult_Helper) CoCreateInstance(
			CLSID_TF_InputProcessorProfiles,
			nullptr,
			CLSCTX_INPROC_SERVER,
			IID_PPV_ARGS(&inputProcProfiles));
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}

		void* tempPtr = nullptr;
		hr = (HResult_Helper) inputProcProfiles->QueryInterface(IID_ITfInputProcessorProfileMgr, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		inputProcProfilesMgr = (ITfInputProcessorProfileMgr*) tempPtr;
	}


	ITfThreadMgr* threadMgr = nullptr;
	ITfThreadMgr2* threadMgr2 = nullptr;
	ITfSourceSingle* threadMgrSource = nullptr;
	{
		auto hr = (HResult_Helper) CoCreateInstance(
			CLSID_TF_ThreadMgr,
			nullptr,
			CLSCTX_INPROC_SERVER,
			IID_PPV_ARGS(&threadMgr));
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		void* tempPtr = nullptr;
		hr = (HResult_Helper) threadMgr->QueryInterface(IID_ITfThreadMgr2, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		threadMgr2 = (ITfThreadMgr2*) tempPtr;

		hr = (HResult_Helper) threadMgr->QueryInterface(IID_ITfSourceSingle, &tempPtr);
		if (hr != HResult_Helper::Ok) {
			std::abort();
		}
		threadMgrSource = (ITfSourceSingle*) tempPtr;
	}

	TfClientId clientId = {};
	hr = (HResult_Helper) threadMgr2->Activate(&clientId);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}


	ITfDocumentMgr* docMgr = nullptr;
	hr = (HResult_Helper) threadMgr2->CreateDocumentMgr(&docMgr);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}

	g_textStore = new tsf_example::TextStoreTest;
	g_textStore->hwnd = hwnd;

	ITfContext* outCtx = nullptr;
	TfEditCookie outTfEditCookie = {};
	hr = (HResult_Helper) docMgr->CreateContext(
		clientId,
		0,
		(ITextStoreACP2*) g_textStore,
		&outCtx,
		&outTfEditCookie);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}

	void* ctxSourceTemp = nullptr;
	hr = (HResult_Helper) outCtx->QueryInterface(IID_ITfSource, &ctxSourceTemp);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}
	auto ctxSource = (ITfSource*) ctxSourceTemp;

	DWORD adviseSinkIdentifier = {};
	hr = (HResult_Helper) ctxSource->AdviseSink(IID_ITfTextEditSink, (ITfTextEditSink*) g_textStore,
												&adviseSinkIdentifier);
	//hr = (HResult_Helper)ctxSource->AdviseSingleSink(clientId, IID_ITfTextEditSink, (ITfTextEditSink*)g_textStore);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}

	hr = (HResult_Helper) docMgr->Push(outCtx);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}

	/*
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
    */

	/*
    hr = (HResult_Helper)threadMgr2->SetFocus(docMgr);
    if (hr != HResult_Helper::Ok) {
        std::abort();
    }
    */

	ITfDocumentMgr* oldDocMgr = nullptr;
	hr = (HResult_Helper) threadMgr->AssociateFocus(hwnd, docMgr, &oldDocMgr);
	if (hr != HResult_Helper::Ok) {
		std::abort();
	}


	// Initialize UI Automation stuff
	auto uiAutomationCoreLib = LoadLibraryA("UIAutomationCore.dll");
	if (uiAutomationCoreLib == nullptr) {
		std::abort();
	}
	auto uiaReturnRawElementProviderFn = GetProcAddress(uiAutomationCoreLib, "UiaReturnRawElementProvider");
	if (uiaReturnRawElementProviderFn == nullptr) {
		std::abort();
	}
	auto uiaHostProviderFromHwndFn = GetProcAddress(uiAutomationCoreLib, "UiaHostProviderFromHwnd");
	if (uiaHostProviderFromHwndFn == nullptr) {
		std::abort();
	}
	tsf_example::uiAutomationFnPtrs.UiaReturnRawElementProvider = (tsf_example::UiaReturnRawElementProvider_FnT)uiaReturnRawElementProviderFn;
	tsf_example::uiAutomationFnPtrs.UiaHostProviderFromHwnd = (tsf_example::UiaHostProviderFromHwnd_FnT)uiaHostProviderFromHwndFn;

	tsf_example::uiAccessProvider = new tsf_example::MyUiAccessProvider();
	tsf_example::uiAccessProvider->hwnd = hwnd;



	MSG msg = {};
	while (true) {
		// We always wait for a new result
		bool messageReceived = GetMessageA(&msg, nullptr, 0, 0);
		if (messageReceived) {
			if ((CustomMessageEnum) msg.message == CustomMessageEnum::CustomMessage) {

				auto job = g_queuedFns.front();
				g_queuedFns.erase(g_queuedFns.begin());
				job();

			} else {
				// We got a message
				TranslateMessage(&msg);
				DispatchMessageA(&msg);
			}
		} else {
			// Error?
		}
	}

	return 0;

}