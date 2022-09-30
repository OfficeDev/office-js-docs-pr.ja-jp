---
title: Office アドイン用語集
description: Office アドインのドキュメント全体で一般的に使用される用語の用語集。
ms.date: 09/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: ef8df6e344698f7d67ebe7afe1759e13630b385d
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234915"
---
# <a name="office-add-ins-glossary"></a>Office アドインの用語集

これは、Office アドインのドキュメント全体で一般的に使用される用語の用語集です。

## <a name="add-in"></a>アドイン

Office アドインは、Office アプリケーションを拡張する Web アプリケーションです。 これらの Web アプリケーションは、外部データの取り込み、プロセスの自動化、Office ドキュメントへの対話型オブジェクトの埋め込みなど、Office アプリケーションに新しい機能を追加します。

Office アドインは VBA、COM、VSTO アドインとは異なります。これは、クロスプラットフォームサポート (通常は Web、Windows、Mac、iPad) を提供し、標準の Web テクノロジ (HTML、CSS、JavaScript) に基づいているためです。 Office アドインの主要なプログラミング言語は、JavaScript または TypeScript です。

## <a name="add-in-commands"></a>アドイン コマンド

**アドイン コマンド** は、アドインの Office UI を拡張する UI 要素 (ボタンやメニューなど) です。 ユーザーがアドイン コマンド要素を選択すると、JavaScript コードの実行や作業ウィンドウへのアドインの表示などのアクションが開始されます。 アドイン コマンドを使用すると、アドインは Office の一部のように見え、アドインに対するユーザーの信頼度が高くなります。 詳細については、「 [Excel、PowerPoint、Word のアドイン コマンドと](../design/add-in-commands.md) [Outlook 用アドイン コマンド](../outlook/add-in-commands-for-outlook.md) 」を参照してください。

リボン [、リボン ボタン](#ribbon-ribbon-button)も参照してください。

## <a name="application"></a>アプリケーション

**アプリケーション** とは、Office アプリケーションを指します。 Office アドインをサポートする Office アプリケーションは、Excel、OneNote、Outlook、PowerPoint、Project、Word です。

クライアント[、](#client)[ホスト](#host)、[Office アプリケーション、Office クライアント](#office-application-office-client)も参照してください。

## <a name="application-specific-api"></a>アプリケーション固有の API

アプリケーション固有の API は、特定の Office アプリケーションにネイティブなオブジェクトと対話する厳密に型指定されたオブジェクトを提供します。 たとえば、Excel JavaScript API を呼び出して、ワークシート、範囲、テーブル、グラフなどにアクセスします。 現在、アプリケーション固有の API は、Excel、OneNote、PowerPoint、Visio、Word で使用できます。 詳細については、 [アプリケーション固有の API モデル](../develop/application-specific-api-model.md) に関するページを参照してください。

[共通 API](#common-api) も参照してください。

## <a name="client"></a>クライアント

**クライアント** は通常、Office アプリケーションを参照します。 Office アドインをサポートする Office アプリケーションまたはクライアントは、Excel、OneNote、Outlook、PowerPoint、Project、Word です。

[アプリケーション](#application)、[ホスト](#host)、[Office アプリケーション、Office クライアント](#office-application-office-client)も参照してください。

## <a name="common-api"></a>Common API

一般的な API は、複数の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスするために使用されます。 この API モデルでは [コールバック](https://developer.mozilla.org/docs/Glossary/Callback_function) が使用され、Office アプリケーションに送信する各要求で 1 つの操作のみを指定できます。

一般的な API は Office 2013 で導入され、Office 2013 以降との対話に使用されます。 一部の一般的な API は、2010 年代初頭のレガシ API です。 Excel、PowerPoint、Word には共通の API 機能がありますが、この機能のほとんどは、アプリケーション固有の API モデルによって置き換えられたり置き換えられたりしています。 可能な場合は、アプリケーション固有の API をお勧めします。

Outlook、UI、認証に関連する一般的な API など、他の一般的な API は、これらの目的に適した最新の API です。 Common API オブジェクト モデルの詳細については、「 [Common JavaScript API オブジェクト モデル](../develop/office-javascript-api-object-model.md)」を参照してください。

[アプリケーション固有の API](#application-specific-api) も参照してください。

## <a name="content-add-in"></a>コンテンツ アドイン

**コンテンツ アドイン** は、Excel、OneNote、または PowerPoint ドキュメントに直接埋め込まれる Web ビューまたは Web ブラウザー ビューです。 コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。 詳細については、「 [Content Office アドイン](../design/content-add-ins.md) 」を参照してください。

[webview](#webview) も参照してください。

## <a name="content-delivery-network-cdn"></a>コンテンツ配信ネットワーク (CDN)

**コンテンツ配信ネットワーク** または **CDN** は、サーバーとデータ センターの分散ネットワークです。 通常、単一のサーバーまたはデータ センターと比較して、リソースの可用性とパフォーマンスが向上します。

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (Contoso と Contoso University とも呼ばれます) は、Microsoft が会社とドメインの例として使用する架空の会社です。

## <a name="custom-function"></a>カスタム関数

**カスタム関数** は、Excel アドインと共にパッケージ化されたユーザー定義関数です。 カスタム関数を使用すると、開発者は、アドインの一部として JavaScript でこれらの関数を定義することで、一般的な Excel 機能を超えて新しい関数を追加できます。 Excel 内のユーザーは、Excel のネイティブ関数と同様にカスタム関数にアクセスできます。 詳細については、「 [Excel でカスタム関数を作成](../excel/custom-functions-overview.md) する」を参照してください。

## <a name="custom-functions-runtime"></a>カスタム関数ランタイム

**カスタム関数ランタイム** は、Office ホストとプラットフォームのいくつかの組み合わせでカスタム関数を実行する [JavaScript 専用ランタイム](../testing/runtimes.md#javascript-only-runtime)です。 UI はなく、Office.js API と対話することはできません。 アドインにカスタム関数しかない場合は、軽量のランタイムを使用することをお勧めします。 カスタム関数が作業ウィンドウまたは Office.js API と対話する必要がある場合は、 [共有ランタイム](../testing/runtimes.md#shared-runtime)を構成します。 詳細については、「 [共有ランタイムを使用するように Office アドインを構成](../develop/configure-your-add-in-to-use-a-shared-runtime.md) する」を参照してください。

[ランタイム](#runtime)、[共有ランタイム](#shared-runtime)も参照してください。

## <a name="custom-functions-only-add-in"></a>カスタム関数専用アドイン

カスタム関数を含むが、作業ウィンドウなどの UI を含まないアドイン。 この種のアドインのカスタム関数は [、JavaScript 専用ランタイム](../testing/runtimes.md#javascript-only-runtime)で実行されます。 UI を含むカスタム関数では、共有ランタイムまたは JavaScript 専用ランタイムと HTML をサポートするランタイムの組み合わせを使用できます。 UI がある場合は、共有ランタイムを使用することをお勧めします。

[カスタム関数](#custom-function)、[カスタム関数ランタイム](#custom-functions-runtime)も参照してください。

## <a name="host"></a>host

**\<Host\>** 通常、Office アプリケーションを参照します。 Office アドインをサポートする Office アプリケーションまたはホストは、Excel、OneNote、Outlook、PowerPoint、Project、Word です。

[アプリケーション](#application)、[クライアント](#client)、[Office アプリケーション、Office クライアント](#office-application-office-client)も参照してください。

## <a name="office-application-office-client"></a>Office アプリケーション、Office クライアント

**Office クライアント** は、Office アプリケーションを参照します。 Office アドインをサポートする Office アプリケーションまたはクライアントは、Excel、OneNote、Outlook、PowerPoint、Project、Word です。

[アプリケーション](#application)、[クライアント](#client)、[ホスト](#host)も参照してください。

## <a name="perpetual"></a>永久

**永続** とは、ボリューム ライセンス契約または小売チャネルを通じて利用可能な Office のバージョンを指します。

その他の Microsoft コンテンツでは、この概念を表すために **サブスクリプション以外** の用語を使用する場合があります。

関連項目: [リテール、リテール永続](#retail-retail-perpetual)、 [ボリューム ライセンス、ボリューム ライセンス付きパーペチュアル、ボリューム ライセンス](#volume-licensed-volume-licensed-perpetual-volume-licensing)

## <a name="platform"></a>platform

**プラットフォーム** とは、通常、Office アプリケーションを実行しているオペレーティング システムを指します。 Office アドインをサポートするプラットフォームには、Windows、Mac、iPad、Web ブラウザーが含まれます。

## <a name="quick-start"></a>クイック スタート

**クイック スタート** は、特定のプログラムの基本的な操作に必要な主要なスキルと知識の概要です。 Office アドインのドキュメントでは、クイック スタートは、Outlook などの特定のアプリケーション用のアドインの開発の概要です。 クイック スタートには、アドイン開発者が約 5 分で完了できる一連の手順が含まれており、その結果、アドインおよび機能開発環境が機能します。

[チュートリアル](#tutorial)も参照してください。

## <a name="requirement-set"></a>要件セット

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="retail-retail-perpetual"></a>リテール、リテール永続

**リテール** とは、リテール チャネルを通じて利用できる永続的なバージョンの Office を指します。 これには、Microsoft 365 サブスクリプションまたはボリューム ライセンス契約によって提供されるバージョンは含まれません。

その他の Microsoft コンテンツでは、この概念を表すために **、1 回限りの購入** または **コンシューマー** という用語を使用する場合があります。

「[永続」](#perpetual)も参照してください。

## <a name="ribbon-ribbon-button"></a>リボン、リボン ボタン

**リボン** は、アプリケーションの機能をウィンドウの上部にある一連のタブまたはボタンに整理するコマンド バーです。 **リボン ボタン** は、このシリーズ内のボタンの 1 つです。 詳細については、「 [Office でリボンを表示または非表示にする](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) 」を参照してください。

## <a name="runtime"></a>ランタイム

**ランタイム** は、アドインが実行されるホスト環境 (JavaScript エンジンと通常は HTML レンダリング エンジンも含む) です。 Office on Windows および Office on Mac では、ランタイムは、Internet Explorer、Edge Legacy、Edge WebView2、Safari などの埋め込みブラウザー コントロール (または Webview) です。 アドインのさまざまな部分は、個別のランタイムで実行されます。 たとえば、アドイン コマンド、カスタム関数、作業ウィンドウ コードは、共有ランタイムを構成しない限り、通常は個別のランタイムを使用 [します](../testing/runtimes.md#shared-runtime)。 詳細については、「 [Office アドインのランタイム](../testing/runtimes.md) と Office アドイン [で使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md) 」を参照してください。

[カスタム関数ランタイム](#custom-functions-runtime)、[共有ランタイム](#shared-runtime)、[Webview](#webview) も参照してください。

## <a name="shared-runtime"></a>共有ランタイム

**共有ランタイム** を使用すると、作業ウィンドウ、アドイン コマンド、カスタム関数など、アドイン内のすべてのコードを同じランタイムで実行し、作業ウィンドウが閉じられた場合でも引き続き実行できます。 詳細については、「 [共有ランタイム](../testing/runtimes.md#shared-runtime) 」と [「Office アドインで共有ランタイムを使用するためのヒント](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) 」を参照してください。

[カスタム関数ランタイム](#custom-functions-runtime)、[ランタイム](#runtime)も参照してください。

## <a name="subscription"></a>サブスクリプション

**サブスクリプション** とは、Microsoft 365 サブスクリプションで使用できる Office のバージョンを指します。

## <a name="task-pane"></a>作業ウィンドウ

作業ウィンドウは、通常、Excel、Outlook、PowerPoint、Word 内のウィンドウの右側に表示されるインターフェイス サーフェイスまたは Web ビューです。 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 ドキュメントに直接機能を埋め込む必要がない場合や埋め込めない場合は、作業ウィンドウを使用します。 詳細については、「 [Office アドインの作業ウィンドウ](../design/task-pane-add-ins.md) 」を参照してください。

[webview](#webview) も参照してください。

## <a name="tutorial"></a>チュートリアル

**チュートリアル** は、ユーザーが製品や手順の使用を学ぶのに役立つ教育支援です。 Office アドインコンテキストでは、チュートリアルでは、Excel などの特定のアプリケーションの完全なアドイン開発プロセスをアドイン開発者に説明します。 これには、次の 20 以上の手順が含まれており、 [クイック スタート](#quick-start)よりも時間の投資が大きくなります。

[クイック スタート](#quick-start)も参照してください。

## <a name="volume-licensed-volume-licensed-perpetual-volume-licensing"></a>ボリューム ライセンス、ボリューム ライセンスの永続、ボリューム ライセンス

**ボリューム ライセンスとは** 、Microsoft と会社間のボリューム ライセンス契約を通じて利用可能な永続的なバージョンの Office を指します。

その他の Microsoft コンテンツでは、この概念を表すために **商用** という用語を使用する場合があります。

「[永続」](#perpetual)も参照してください。

## <a name="web-add-in"></a>Web アドイン

**Web アドイン** は、Office アドインのレガシ用語です。 この用語は、Microsoft 365 ドキュメントで最新の Office アドインを VBA、COM、VSTO などの他の種類のアドインと区別する必要がある場合に使用できます。

アドインも[参照してください。](#add-in)

## <a name="webview"></a>Webview

**Web ビュー** は、アプリケーション内に Web コンテンツを表示する要素またはビューです。 コンテンツ アドインと作業ウィンドウの両方に埋め込み Web ブラウザーが含まれており、Office アドインの Web ビューの例です。

[コンテンツ アドイン](#content-add-in)、[作業ウィンドウ](#task-pane)も参照してください。

## <a name="xll"></a>XLL

**XLL** アドインは、ユーザー定義関数を提供し、ファイル拡張子 **.xll** を持つ Excel アドイン ファイルです。 XLL ファイルは、Excel でのみ開くことができるダイナミック リンク ライブラリ (DLL) ファイルの一種です。 XLL アドイン ファイルは、C または C++ で記述する必要があります。 カスタム関数は、XLL ユーザー定義関数に相当する最新の関数です。 カスタム関数は、プラットフォーム間のサポートを提供し、XLL ファイルと下位互換性があります。 詳細については、「 [XLL ユーザー定義関数を使用してカスタム関数を拡張](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) する」を参照してください。

[カスタム関数](#custom-function)も参照してください。

## <a name="yeoman-generator-yo-office"></a>Yeoman ジェネレーター、 yo office

[Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)は、オープンソース [Yeoman](https://github.com/yeoman/yo) ツールを使用して、コマンド ラインを使用して Office アドインを生成します。 `yo office` は、Office アドイン用の Yeoman ジェネレーターを実行するコマンドです。Office アドインのクイック スタートとチュートリアルでは、Yeoman ジェネレーターを使用します。

## <a name="see-also"></a>関連項目

- [Office アドインのその他のリソース](resources-links-help.md)