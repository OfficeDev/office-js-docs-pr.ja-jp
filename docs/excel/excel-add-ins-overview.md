# <a name="excel-add-ins-overview"></a>Excel アドインの概要

Excel アドインを使用すると、Office for Windows、Office Online、Office for the Mac、および Office for the iPad など、複数のプラットフォームにわたって Excel アプリケーションの機能を拡張できます。 ブック内で Excel アドインを使用すると、次の操作を実行できます。

- Excel オブジェクトを操作して Excel データを読み書きします 
- Web ベースの作業ウィンドウまたはコンテンツ ウィンドウを使用して機能を拡張します 
- カスタム リボン ボタンやコンテキスト メニューの項目を追加します
- ダイアログ ウィンドウを使用して充実した操作を提供します 

Office アドインのプラットフォームには、Excel アドインの作成と実行を可能にするフレームワークと Office.js JavaScript API が用意されています。Office アドインのプラットフォームを使用した Excel アドインの作成には、次の利点があります。

* **クロスプラットフォーム サポート**:Excel アドインは、Windows 版、Mac 版、iOS 版の Office と、Office Online で実行できます。
* **一元展開**:管理者は、組織全体のユーザーに Excel アドインをすばやく簡単に展開できます。
* **シングル サインオン (SSO)**:Excel のアドインを Microsoft Graph に簡単に統合できます。
* **標準の Web テクノロジの使用**:HTML、CSS、JavaScript などの一般的な Web テクノロジを使用する Excel アドインを作成します。
* **Office ストアを経由した配布**:Excel アドインを [Office ストア](https://store.office.com/en-us/appshome.aspx)に公開することで、幅広いユーザーと共有します。

> **注**:Excel アドインは、Office for Windows 上でのみ実行する以前の Office 統合ソリューションである COM アドインや VSTO アドインとは異なります。 COM アドインとは異なり、Excel アドインではユーザーのデバイスや Excel 内にコードをインストールする必要はありません。 

## <a name="components-of-an-excel-add-in"></a>Excel アドインのコンポーネント 

Excel アドインには 2 つの基本コンポーネントが含まれています。Web アプリケーションと、マニフェスト ファイルと呼ばれる構成ファイルです。 

Web アプリケーションは、[JavaScript API for Office](../../reference/javascript-api-for-office.md) を使用して Excel のオブジェクトを操作します。また、オンライン リソースとの相互操作を簡単にすることもできます。 たとえば、アドインでは次の操作を実行できます。

* ブック内のデータ (ワークシート、範囲、表、グラフ、名前付きの項目など) を作成、読み込み、更新、および削除します。
* 標準の OAuth 2.0 のフローを使用して、オンライン サービスでユーザー認証を実行します。
* Microsoft Graph やその他の API に、API 要求を発行します。

Web アプリケーションは、任意の Web サーバー上でホストできます。また、クライアント側のフレームワーク (Angular、React、jQuery など) や、サーバー側のテクノロジ (ASP.NET、Node.js、PHP など) を使用して構築できます。

[マニフェスト](../overview/add-in-manifests.md)は XML 構成ファイルであり、次のような設定と機能を指定することによって、アドインと Office クライアントを統合する方法を定義します。 

* アドインの Web アプリケーションの URL。
* アドインの表示名、説明、ID、バージョン、および既定のロケール。
* アドインと Excel を統合する方法。アドインが作成する任意のカスタム UI (リボンのボタン、コンテキスト メニューなど) の統合を含む。
* ドキュメントの読み取り、書き込みなど、アドインに必要なアクセス許可。

エンドユーザーが Excel アドインをインストールして使用できるようにするには、そのマニフェストを Office ストアかアドイン カタログに公開する必要があります。 

## <a name="capabilities-of-an-excel-add-in"></a>Excel アドインの機能

ブック内のコンテンツの操作の他に、Excel アドインでは、カスタム リボンのボタンやメニュー コマンドを追加したり、作業ウィンドウを挿入したり、ダイアログ ボックスを開いたり、グラフや対話型のビジュアル化などの豊富な Web ベースのオブジェクトをワークシート内に埋め込むことができます (次のスクリーン ショットを参照)。 これらの各機能の詳細については、「[Excel の機能を拡張する](excel-add-ins-extend-excel.md)」を参照してください。

**カスタム リボンのボタン**

![アドイン コマンド](../../images/Excel_add-in_commands_Script-Lab.png)

**作業ウィンドウ**

![アドイン作業ウィンドウ](../../images/Excel_add-in_task_pane_Insights.png)

**ダイアログ ボックス**

![アドイン ダイアログ ボックス](../../images/Excel_add-in_dialog_choose-number.png)

**コンテンツ アドイン**

![コンテンツ アドイン](../../images/Excel_add-in_content_map.png)

## <a name="javascript-apis-to-interact-with-workbook-content"></a>ブックのコンテンツを操作する JavaScript API

Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む [JavaScript API for Office](../../reference/javascript-api-for-office.md) を使用して、Excel のオブジェクトを操作します。

* **Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](../../reference/excel/excel-add-ins-reference-overview.md) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定された Excel オブジェクトが用意されています。 

* **Shared API**:Office 2013 で導入された共有 API を使用すると、Word、Excel、PowerPoint など複数の種類のホスト アプリケーションに共通する UI、ダイアログ、クライアント設定などの機能にアクセスできます。 共有 API は Excel の操作に限られた機能を提供します。そのため、アドインを Excel 2013 で実行する必要がある場合に使用できます。

## <a name="next-steps"></a>次の手順

[最初の Excel アドインを作成する](excel-add-ins-get-started-overview.md)ことから始めます。 次に、Excel アドイン構築の[中心概念](excel-add-ins-core-concepts.md)について説明します。

## <a name="additional-resources"></a>その他のリソース

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドイン開発のベスト プラクティス](../overview/add-in-development-best-practices.md)
- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API リファレンス](../../reference/excel/excel-add-ins-reference-overview.md)
