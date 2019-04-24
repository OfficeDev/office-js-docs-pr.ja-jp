---
title: Windows 10 で F12 開発者ツールを使用してアドインをデバッグする
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 750411bea187a0ade9b3723e3198d82f7c482c9f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450151"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a>Windows 10 で F12 開発者ツールを使用してアドインをデバッグする

Windows 10 に含まれている F12 開発者ツールにより、web ページのデバッグ、テスト、および高速化ができます。 それらを使用すれば、Visual Studio などの IDE を使用していない場合や、アドインを IDE の外部で実行中に問題を調査する必要がある場合に、Office アドインの開発とデバッグを行うこともできます。 この記事では、Windows 10 で F12 開発者ツールのデバッガー ツールを使用して、ご利用の Office アドインをテストする方法について説明します。

> [!NOTE]
> この記事の手順を使用して、実行関数を使用する Outlook アドインをデバッグすることはできません。 実行関数を使用する Outlook アドインのデバッグには、スクリプト モードの Visual Studio またはその他のスクリプト デバッガーにアタッチすることをお勧めします。

## <a name="prerequisites"></a>前提条件

以下のソフトウェアが必要です。

- Windows 10 に含まれる F12 開発者ツール。 
    
- アドインをホストする Office クライアント アプリケーション。  
    
- アドイン。  

## <a name="using-the-debugger"></a>デバッガーの使用

Windows 10 の F12 開発者ツールからデバッガーを使用して、AppSource からのアドインやその他の場所から追加したアドインをテストすることができます。 アドインの実行後、F12 開発者ツールを起動できます。 F12 ツールは個別のウィンドウに表示され、Visual Studio を使用しません。

> [!NOTE]
> デバッガーは、Windows 10 および Internet Explorer 上の F12 開発者ツールの一部です。Windows の以前のバージョンにはデバッガーは含まれません。 

次の例では、AppSource から Word と無料のアドインを使用します。

1. Word を起動し、空白の文書を選択します。 
    
2. アドイン グループの [**挿入**] タブで、[**ストア**]、[**QR4Office**] アドインの順に選択します  (ストアやアドイン カタログから、任意のアドインを読み込むことができます)。
    
3. ご利用の Office のバージョンに対応する F12 開発者ツールを起動します。
    
   - 32 ビット版の Office の場合は、C:\Windows\System32\F12\IEChooser.exe を使用します
    
   - 64 ビット版の Office の場合は、C:\Windows\SysWOW64\F12\IEChooser.exe を使用します
    
   IEChooser を起動すると、[デバッグするターゲットの選択] という名前の別ウィンドウに、デバッグ可能なアプリケーションが表示されます。 関心があるアプリケーションを選択します。 独自のアドインを記述している場合、アドインを展開した Web サイトを選択します。これは、localhost の URL である可能性があります。 
    
   たとえば、**home.html** を選択します。 
    
   ![バブルのアドインをポイントする IEChooser 画面](../images/choose-target-to-debug.png)

4. F12 ウィンドウで、デバッグするファイルを選択します。
    
   F12 ウィンドウのファイルを選択するには、**スクリプト** (左側) ウィンドウの上にあるフォルダー アイコンを選びます。 ドロップダウン リストに表示される利用可能なファイルのリストから [**Home.js**] を選択します。
    
5. ブレークポイントを設定します。
    
   **Home.js** にブレークポイントを設定するために、`textChanged` 関数内の行 144 を選択します。 その行の左側と **[呼び出し履歴] と [ブレークポイント]** (右下) ウィンドウの対応する行に赤い点が表示されます。 ブレークポイントを設定するその他の方法については、「[デバッガーを使用して実行中の JavaScript を検査する](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))」を参照してください。 
    
   ![home.js ファイルのブレーキポイントを含むデバッガー](../images/debugger-home-js-02.png)

6. アドインを実行して、ブレークポイントをトリガーします。
    
   Word で、[**QR4Office**] ウィンドウの上部にある [URL] テキスト ボックスを選択して、テキストを入力してみます。 デバッガー内の **[呼び出し履歴] と [ブレークポイント]** ウィンドウで、ブレークポイントがトリガーされ、さまざまな情報が表示されることがわかります。 結果を確認するには、デバッガーの更新が必要な場合があります。
    
   ![トリガーされたブレークポイントの結果を含むデバッガー](../images/debugger-home-js-01.png)


## <a name="see-also"></a>関連項目

- [デバッガーを使用して実行中の JavaScript を検査する](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- 
  [F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
