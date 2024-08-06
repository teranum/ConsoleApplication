#pragma once
//
#include <atlcomcli.h>

class AxKHOpenApi :
	public IOleClientSite,
	public IOleInPlaceSite,
	public IStorage
{
	CComPtr<IOleObject> _oleObj;

public:
	bool Create(HWND hParent)
	{
		return true;
	}

	// Inherited via IOleClientSite
	HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override { return E_NOTIMPL; }
	ULONG __stdcall AddRef(void) override { return 0; }
	ULONG __stdcall Release(void) override { return 0; }
	HRESULT __stdcall SaveObject(void) override { return E_NOTIMPL; }
	HRESULT __stdcall GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk) override { return E_NOTIMPL; }
	HRESULT __stdcall GetContainer(IOleContainer** ppContainer) override { return E_NOTIMPL; }
	HRESULT __stdcall ShowObject(void) override { return S_OK; }
	HRESULT __stdcall OnShowWindow(BOOL fShow) override { return E_NOTIMPL; }
	HRESULT __stdcall RequestNewObjectLayout(void) override { return E_NOTIMPL; }

	// Inherited via IOleInPlaceSite
	HRESULT __stdcall GetWindow(HWND* phwnd) override { return E_NOTIMPL; }
	HRESULT __stdcall ContextSensitiveHelp(BOOL fEnterMode) override { return E_NOTIMPL; }
	HRESULT __stdcall CanInPlaceActivate(void) override { return E_NOTIMPL; }
	HRESULT __stdcall OnInPlaceActivate(void) override { return S_OK; }
	HRESULT __stdcall OnUIActivate(void) override { return S_OK; }
	HRESULT __stdcall GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) override { return E_NOTIMPL; }
	HRESULT __stdcall Scroll(SIZE scrollExtant) override { return E_NOTIMPL; }
	HRESULT __stdcall OnUIDeactivate(BOOL fUndoable) override { return E_NOTIMPL; }
	HRESULT __stdcall OnInPlaceDeactivate(void) override { return S_OK; }
	HRESULT __stdcall DiscardUndoState(void) override { return E_NOTIMPL; }
	HRESULT __stdcall DeactivateAndUndo(void) override { return E_NOTIMPL; }
	HRESULT __stdcall OnPosRectChange(LPCRECT lprcPosRect) override { return E_NOTIMPL; }

	// Inherited via IStorage
	HRESULT __stdcall CreateStream(const OLECHAR* pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStream** ppstm) override { return E_NOTIMPL; }
	HRESULT __stdcall OpenStream(const OLECHAR* pwcsName, void* reserved1, DWORD grfMode, DWORD reserved2, IStream** ppstm) override { return E_NOTIMPL; }
	HRESULT __stdcall CreateStorage(const OLECHAR* pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStorage** ppstg) override { return E_NOTIMPL; }
	HRESULT __stdcall OpenStorage(const OLECHAR* pwcsName, IStorage* pstgPriority, DWORD grfMode, SNB snbExclude, DWORD reserved, IStorage** ppstg) override { return E_NOTIMPL; }
	HRESULT __stdcall CopyTo(DWORD ciidExclude, const IID* rgiidExclude, SNB snbExclude, IStorage* pstgDest) override { return E_NOTIMPL; }
	HRESULT __stdcall MoveElementTo(const OLECHAR* pwcsName, IStorage* pstgDest, const OLECHAR* pwcsNewName, DWORD grfFlags) override { return E_NOTIMPL; }
	HRESULT __stdcall Commit(DWORD grfCommitFlags) override { return E_NOTIMPL; }
	HRESULT __stdcall Revert(void) override { return E_NOTIMPL; }
	HRESULT __stdcall EnumElements(DWORD reserved1, void* reserved2, DWORD reserved3, IEnumSTATSTG** ppenum) override { return E_NOTIMPL; }
	HRESULT __stdcall DestroyElement(const OLECHAR* pwcsName) override { return E_NOTIMPL; }
	HRESULT __stdcall RenameElement(const OLECHAR* pwcsOldName, const OLECHAR* pwcsNewName) override { return E_NOTIMPL; }
	HRESULT __stdcall SetElementTimes(const OLECHAR* pwcsName, const FILETIME* pctime, const FILETIME* patime, const FILETIME* pmtime) override { return E_NOTIMPL; }
	HRESULT __stdcall SetClass(REFCLSID clsid) override { return S_OK; }
	HRESULT __stdcall SetStateBits(DWORD grfStateBits, DWORD grfMask) override { return E_NOTIMPL; }
	HRESULT __stdcall Stat(STATSTG* pstatstg, DWORD grfStatFlag) override { return E_NOTIMPL; }
};

