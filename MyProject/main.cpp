// PowerPoint 撮影報告書 自動生成ツール（ダイアログ GUI 版）
// COM オートメーションを使用して見本テンプレートから新規報告書を生成する

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#undef GetProp  // winuser.h が GetProp を GetPropW にマクロ定義するため Disp::GetProp と衝突する
#include <comdef.h>
#include <oleauto.h>
#include <atlbase.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <sstream>
#include <stdexcept>
#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")

// ─── 報告書データ ───────────────────────────────────────────────────────────

struct ReportData
{
    std::wstring customerName  = L"（株）ニッコー様";
    std::wstring reportTitle   = L"撮影報告書";
    std::wstring shootingDate  = L"2026年　月　日";
    std::wstring issuerName    = L"コムスキャンテクノ株式会社（JOHNAN Group）";

    std::wstring purposeText =
        L"撮影目的は、ホーンのSampleA～SampleCをＣＴ撮影を行って、\n"
        L"内部コイルと鉄心の位置関係等を計測する。";

    std::wstring overviewText =
        L"サンプルの撮影風景を以下に掲載します。\n"
        L"・最大管電圧 = 150kVのCT装置を使用して撮影。\n"
        L"・銅フィルターを用いてノイズを低減。\n"
        L"・積算平均を増やすことでS/Nをアップし、ノイズを低減。";

    std::wstring conditionText = L"以下に撮影条件を記載します。";

    std::wstring processText =
        L"以下に解析プロセスの流れと各プロセス概要説明を記載します。";

    std::wstring resultText =
        L"・内部コイルや鉄心などの部品を明確に把握できるようCT撮影を実施できました。\n"
        L"・鉄心の隙間計測、電磁コイル等の部品配置を確認することは可能である。\n"
        L"・SampleCにおいては、電磁コイルが実装されていないことを確認しています。\n"
        L"\n"
        L"・以下のポイントを計測しました。結果を下記表に記載します。\n"
        L"※次ページに撮影画像を掲載。\n"
        L"\n"
        L"・Ｘ線管：最大管電圧 = 150kV";

    struct Sample {
        std::wstring name;
        std::wstring imagePath;
        std::wstring imagePathMeasure;
        std::wstring note;
    };
    std::vector<Sample> samples = {
        { L"Sample1", L"", L"", L"" },
        { L"Sample2", L"", L"", L"" },
        { L"Sample3", L"", L"", L"電磁コイルが実装されていないことを確認。" },
    };
};

// ─── COM IDispatch ヘルパー ─────────────────────────────────────────────────

class Disp
{
    CComPtr<IDispatch> m_p;

    DISPID GetID(const wchar_t* name) const {
        if (!m_p) return DISPID_UNKNOWN;
        DISPID d = DISPID_UNKNOWN;
        BSTR b = SysAllocString(name);
        m_p->GetIDsOfNames(IID_NULL, &b, 1, LOCALE_USER_DEFAULT, &d);
        SysFreeString(b);
        return d;
    }

public:
    Disp() = default;
    Disp(IDispatch* p) : m_p(p) {}
    explicit operator bool() const { return m_p != nullptr; }

    _variant_t GetProp(const wchar_t* name) const {
        _variant_t res;
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return res;
        DISPPARAMS dp{};
        m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_PROPERTYGET, &dp, &res, nullptr, nullptr);
        return res;
    }

    _variant_t GetPropIdx(const wchar_t* name, long idx) const {
        _variant_t res;
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return res;
        _variant_t vi(idx);
        VARIANT args[1]{ vi.Detach() };
        DISPPARAMS dp{ args, nullptr, 1, 0 };
        m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_PROPERTYGET, &dp, &res, nullptr, nullptr);
        return res;
    }

    HRESULT SetProp(const wchar_t* name, _variant_t val) const {
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return E_FAIL;
        DISPID put = DISPID_PROPERTYPUT;
        VARIANT v = val.Detach();
        DISPPARAMS dp{ &v, &put, 1, 1 };
        return m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                           DISPATCH_PROPERTYPUT, &dp, nullptr, nullptr, nullptr);
    }

    _variant_t Call(const wchar_t* name) const {
        _variant_t res;
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return res;
        DISPPARAMS dp{};
        m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &dp, &res, nullptr, nullptr);
        return res;
    }

    _variant_t Call1(const wchar_t* name, _variant_t a0) const {
        _variant_t res;
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return res;
        VARIANT args[1]{ a0.Detach() };
        DISPPARAMS dp{ args, nullptr, 1, 0 };
        m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &dp, &res, nullptr, nullptr);
        return res;
    }

    _variant_t Call2(const wchar_t* name, _variant_t a0, _variant_t a1) const {
        _variant_t res;
        DISPID id = GetID(name);
        if (id == DISPID_UNKNOWN) return res;
        VARIANT args[2]{ a1.Detach(), a0.Detach() }; // COM は引数逆順
        DISPPARAMS dp{ args, nullptr, 2, 0 };
        m_p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD, &dp, &res, nullptr, nullptr);
        return res;
    }

    Disp Child(const wchar_t* name) const {
        _variant_t v = GetProp(name);
        if (V_VT(&v) == VT_DISPATCH && V_DISPATCH(&v))
            return Disp(V_DISPATCH(&v));
        return {};
    }

    Disp ChildIdx(const wchar_t* name, long idx) const {
        _variant_t v = GetPropIdx(name, idx);
        if (V_VT(&v) == VT_DISPATCH && V_DISPATCH(&v))
            return Disp(V_DISPATCH(&v));
        return {};
    }

    long LongProp(const wchar_t* name) const {
        _variant_t v = GetProp(name);
        if (V_VT(&v) == VT_I4) return V_I4(&v);
        if (V_VT(&v) == VT_I2) return (long)V_I2(&v);
        if (V_VT(&v) == VT_R8) return (long)V_R8(&v);
        return 0;
    }

    std::wstring StrProp(const wchar_t* name) const {
        _variant_t v = GetProp(name);
        if (V_VT(&v) == VT_BSTR && V_BSTR(&v))
            return V_BSTR(&v);
        return {};
    }
};

// ─── ユーティリティ ─────────────────────────────────────────────────────────

static void SetTextByName(Disp slide, const std::wstring& shapeName, const std::wstring& text)
{
    Disp shapes = slide.Child(L"Shapes");
    if (!shapes) return;
    long n = shapes.LongProp(L"Count");
    for (long i = 1; i <= n; i++) {
        Disp sh = shapes.ChildIdx(L"Item", i);
        if (!sh) continue;
        if (sh.StrProp(L"Name") == shapeName) {
            Disp tf = sh.Child(L"TextFrame");
            if (!tf) return;
            Disp tr = tf.Child(L"TextRange");
            if (!tr) return;
            tr.SetProp(L"Text", _variant_t(text.c_str()));
            return;
        }
    }
}

// ─── ダイアログ状態 ─────────────────────────────────────────────────────────

struct DlgState {
    HWND hDlg = nullptr;
    bool converting = false;
};

static DlgState g_state;

static void Log(const std::wstring& msg)
{
    HWND hLog = GetDlgItem(g_state.hDlg, IDC_LOG);
    int len = GetWindowTextLengthW(hLog);
    SendMessageW(hLog, EM_SETSEL, len, len);
    std::wstring line = msg + L"\r\n";
    SendMessageW(hLog, EM_REPLACESEL, FALSE, (LPARAM)line.c_str());
}

// ─── コンバート処理 ─────────────────────────────────────────────────────────

static bool RunConvert(const std::wstring& templatePath, const std::wstring& outputPath)
{
    ReportData data;

    Log(L"PowerPoint を起動中...");

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"PowerPoint.Application", &clsid);
    if (FAILED(hr)) { Log(L"エラー: PowerPoint が見つかりません"); return false; }

    CComPtr<IUnknown> pUnk;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IUnknown, (void**)&pUnk);
    if (FAILED(hr)) { Log(L"エラー: PowerPoint 起動失敗"); return false; }

    CComPtr<IDispatch> pAppRaw;
    hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pAppRaw);
    if (FAILED(hr)) { Log(L"エラー: IDispatch 取得失敗"); return false; }

    Disp app(pAppRaw);
    app.SetProp(L"Visible", _variant_t(false));

    Log(L"テンプレートを開いています: " + templatePath);

    Disp presentations = app.Child(L"Presentations");
    if (!presentations) { Log(L"エラー: Presentations 取得失敗"); app.Call(L"Quit"); return false; }

    _variant_t vPres = presentations.Call2(L"Open",
        _variant_t(templatePath.c_str()), _variant_t(false));
    if (V_VT(&vPres) != VT_DISPATCH || !V_DISPATCH(&vPres)) {
        Log(L"エラー: テンプレートを開けませんでした");
        app.Call(L"Quit");
        return false;
    }

    Disp pres(V_DISPATCH(&vPres));
    Disp slides = pres.Child(L"Slides");
    long slideCount = slides.LongProp(L"Count");
    Log(L"スライド数: " + std::to_wstring(slideCount));

    // スライド1: タイトル
    if (slideCount >= 1) {
        Disp s = slides.ChildIdx(L"Item", 1);
        if (s) {
            SetTextByName(s, L"顧客名",         data.customerName);
            SetTextByName(s, L"報告書タイトル",  data.reportTitle);
            SetTextByName(s, L"撮影日付",        data.shootingDate);
            SetTextByName(s, L"発行会社",        data.issuerName);
            Log(L"スライド1 (タイトル) 更新完了");
        }
    }

    // スライド2: 撮影目的
    if (slideCount >= 2) {
        Disp s = slides.ChildIdx(L"Item", 2);
        if (s) { SetTextByName(s, L"本文", data.purposeText); Log(L"スライド2 (撮影目的) 更新完了"); }
    }

    // スライド3: 撮影概要
    if (slideCount >= 3) {
        Disp s = slides.ChildIdx(L"Item", 3);
        if (s) { SetTextByName(s, L"本文", data.overviewText); Log(L"スライド3 (撮影概要) 更新完了"); }
    }

    // スライド4: 撮影条件
    if (slideCount >= 4) {
        Disp s = slides.ChildIdx(L"Item", 4);
        if (s) { SetTextByName(s, L"本文", data.conditionText); Log(L"スライド4 (撮影条件) 更新完了"); }
    }

    // スライド5: 解析プロセス
    if (slideCount >= 5) {
        Disp s = slides.ChildIdx(L"Item", 5);
        if (s) { SetTextByName(s, L"本文", data.processText); Log(L"スライド5 (解析プロセス) 更新完了"); }
    }

    // スライド6: 結果
    if (slideCount >= 6) {
        Disp s = slides.ChildIdx(L"Item", 6);
        if (s) { SetTextByName(s, L"本文", data.resultText); Log(L"スライド6 (結果) 更新完了"); }
    }

    Log(L"保存中: " + outputPath);
    pres.Call1(L"SaveAs", _variant_t(outputPath.c_str()));
    pres.Call1(L"Close", _variant_t(false));
    app.Call(L"Quit");

    Log(L"完了！ → " + outputPath);
    return true;
}

// ─── ファイル選択ダイアログ ─────────────────────────────────────────────────

static bool BrowseOpenFile(HWND hParent, std::wstring& outPath)
{
    wchar_t buf[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = hParent;
    ofn.lpstrFilter = L"PowerPoint ファイル (*.pptx)\0*.pptx\0すべてのファイル (*.*)\0*.*\0";
    ofn.lpstrFile   = buf;
    ofn.nMaxFile    = MAX_PATH;
    ofn.Flags       = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    if (!GetOpenFileNameW(&ofn)) return false;
    outPath = buf;
    return true;
}

static bool BrowseSaveFile(HWND hParent, std::wstring& outPath)
{
    wchar_t buf[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize     = sizeof(ofn);
    ofn.hwndOwner       = hParent;
    ofn.lpstrFilter     = L"PowerPoint ファイル (*.pptx)\0*.pptx\0すべてのファイル (*.*)\0*.*\0";
    ofn.lpstrFile       = buf;
    ofn.nMaxFile        = MAX_PATH;
    ofn.lpstrDefExt     = L"pptx";
    ofn.Flags           = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    if (!GetSaveFileNameW(&ofn)) return false;
    outPath = buf;
    return true;
}

// ─── ダイアログプロシージャ ─────────────────────────────────────────────────

static INT_PTR CALLBACK DlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM)
{
    switch (msg)
    {
    case WM_INITDIALOG:
    {
        g_state.hDlg = hDlg;

        // デフォルトパスを設定
        wchar_t exeDir[MAX_PATH];
        GetModuleFileNameW(nullptr, exeDir, MAX_PATH);
        std::wstring dir = exeDir;
        auto pos = dir.rfind(L'\\');
        if (pos != std::wstring::npos) dir = dir.substr(0, pos);

        // exeDir から4階層上が見本ファイルのある場所（出力ディレクトリ構成に依存）
        std::wstring templateDefault = dir + L"\\..\\..\\..\\..\\見本_サンプル撮影報告書.pptx";
        std::wstring outputDefault   = dir + L"\\..\\..\\..\\..\\新規_撮影報告書.pptx";

        wchar_t templateAbs[MAX_PATH], outputAbs[MAX_PATH];
        GetFullPathNameW(templateDefault.c_str(), MAX_PATH, templateAbs, nullptr);
        GetFullPathNameW(outputDefault.c_str(),   MAX_PATH, outputAbs,   nullptr);

        SetDlgItemTextW(hDlg, IDC_TEMPLATE_PATH, templateAbs);
        SetDlgItemTextW(hDlg, IDC_OUTPUT_PATH,   outputAbs);
        return TRUE;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_TEMPLATE_BROWSE:
        {
            std::wstring path;
            if (BrowseOpenFile(hDlg, path))
                SetDlgItemTextW(hDlg, IDC_TEMPLATE_PATH, path.c_str());
            break;
        }
        case IDC_OUTPUT_BROWSE:
        {
            std::wstring path;
            if (BrowseSaveFile(hDlg, path))
                SetDlgItemTextW(hDlg, IDC_OUTPUT_PATH, path.c_str());
            break;
        }
        case IDC_CONVERT_BTN:
        {
            if (g_state.converting) break;

            wchar_t tmpl[MAX_PATH] = {}, out[MAX_PATH] = {};
            GetDlgItemTextW(hDlg, IDC_TEMPLATE_PATH, tmpl, MAX_PATH);
            GetDlgItemTextW(hDlg, IDC_OUTPUT_PATH,   out,  MAX_PATH);

            if (tmpl[0] == L'\0') { MessageBoxW(hDlg, L"テンプレートファイルを指定してください。", L"エラー", MB_ICONERROR); break; }
            if (out[0]  == L'\0') { MessageBoxW(hDlg, L"出力先ファイルを指定してください。",         L"エラー", MB_ICONERROR); break; }

            g_state.converting = true;
            EnableWindow(GetDlgItem(hDlg, IDC_CONVERT_BTN), FALSE);
            SetDlgItemTextW(hDlg, IDC_LOG, L"");

            bool ok = false;
            try {
                ok = RunConvert(tmpl, out);
            }
            catch (const _com_error& e) {
                Log(std::wstring(L"COMエラー: ") + e.ErrorMessage());
            }
            catch (const std::exception& e) {
                int len = (int)strlen(e.what());
                std::wstring ws(len, L' ');
                MultiByteToWideChar(CP_ACP, 0, e.what(), len, ws.data(), len);
                Log(L"エラー: " + ws);
            }

            EnableWindow(GetDlgItem(hDlg, IDC_CONVERT_BTN), TRUE);
            g_state.converting = false;

            if (ok) MessageBoxW(hDlg, L"変換が完了しました。", L"完了", MB_ICONINFORMATION);
            break;
        }
        case IDCANCEL:
            EndDialog(hDlg, 0);
            break;
        }
        return TRUE;

    case WM_CLOSE:
        EndDialog(hDlg, 0);
        return TRUE;
    }
    return FALSE;
}

// ─── エントリポイント ────────────────────────────────────────────────────────

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int)
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        MessageBoxW(nullptr, L"COM 初期化失敗", L"エラー", MB_ICONERROR);
        return 1;
    }

    InitCommonControls();

    DialogBoxW(hInst, MAKEINTRESOURCEW(IDD_MAIN_DIALOG), nullptr, DlgProc);

    CoUninitialize();
    return 0;
}
