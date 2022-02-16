---
title: Officeアドイン用語集
description: アドインのドキュメント全体で一般的にOffice用語集。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0c83f056f4eea9c8750bbf4c2d47a2888af96ec2
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855731"
---
# <a name="office-add-ins-glossary"></a>Officeアドイン用語集

これは、アドインのドキュメント全体で一般的に使用Office用語集です。

## <a name="add-in"></a>アドイン

Officeアドインは、アプリケーションを拡張するOfficeです。 これらの Web アプリケーションは、外部データの取り込み、プロセスの自動化、ドキュメントへの対話型オブジェクトの埋め込みなど、Office アプリケーションに新しい機能を追加Officeします。

Officeアドインは、クロスプラットフォーム サポート (通常は web、Windows、Mac、および iPad) を提供し、標準 Web テクノロジ (HTML、CSS、JavaScript) に基づくため、VBA、COM、および VSTO アドインとは異なります。 アドインの主なプログラミング言語Office JavaScript または TypeScript です。

## <a name="add-in-commands"></a>アドイン コマンド

**アドイン コマンドは、** ボタンやメニューなどの UI 要素で、アドインOffice UI を拡張します。 ユーザーがアドイン コマンド要素を選択すると、JavaScript コードの実行や作業ウィンドウでのアドインの表示などのアクションが開始されます。 アドイン コマンドを使用すると、アドインが Office の一部のように見えるので、ユーザーはアドインに対する信頼度が高くなっています。 詳細[については、「Add-in commands for Excel、](../design/add-in-commands.md)PowerPoint」、および「Word およびアドイン [コマンド](../outlook/add-in-commands-for-outlook.md)Outlook参照してください。

「リボン、 [リボン ボタン」も参照してください](#ribbon-ribbon-button)。

## <a name="application"></a>アプリケーション

**アプリケーション** は、アプリケーションをOfficeします。 Office アドインをサポートする Office アプリケーションは、Excel、OneNote、Outlook、PowerPoint、Project、および Word です。

「client、[host](#client)[、Office](#host)[、Office」も参照してください](#office-application-office-client)。

## <a name="application-specific-api"></a>アプリケーション固有の API

アプリケーション固有の API は、特定のアプリケーションにネイティブなオブジェクトを操作する、強力に型指定されたオブジェクトOfficeします。 たとえば、ワークシート、範囲、表、グラフExcelアクセスするために JavaScript API を呼び出します。 現在、アプリケーション固有の API は、Excel、OneNote、PowerPoint、Visio、および Word で使用できます。 詳細 [については、「アプリケーション固有の API モデル](../develop/application-specific-api-model.md) 」を参照してください。

「Common [API」も参照してください](#common-api)。

## <a name="client"></a>クライアント

**クライアント** は通常、アプリケーションをOfficeします。 Office アドインをサポートする Office アプリケーション、またはクライアントは、Excel、OneNote、Outlook、PowerPoint、Project、および Word です。

「application、[host](#application)[、Office](#host)[、Office」も参照してください](#office-application-office-client)。

## <a name="common-api"></a>Common API

一般的な API は、複数のアプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスOfficeされます。 この API モデルでは [コールバック](https://developer.mozilla.org/docs/Glossary/Callback_function) が使用され、Office アプリケーションに送信する各要求で 1 つの操作のみを指定できます。

一般的な API は 2013 年Office導入され、2013 以降のユーザーとの対話Office使用されます。 一部の一般的な API は、2010 の初期の従来の API です。 Excel、PowerPoint、Word はすべて共通の API 機能を持っていますが、ほとんどの機能はアプリケーション固有の API モデルに置き換えられたり置き換えられたりしています。 可能な場合は、アプリケーション固有の API が優先されます。

その他の一般的な API (Outlook、UI、認証に関連する一般的な API など)は、これらの目的に合った最新の API です。 共通 API オブジェクト モデルの詳細については、「 [Common JavaScript API オブジェクト モデル」を参照してください](../develop/office-javascript-api-object-model.md)。

「アプリケーション固有の [API」も参照してください](#application-specific-api)。

## <a name="content-add-in"></a>コンテンツ アドイン

**コンテンツ アドインは、** Web ビューまたは Web ブラウザー ビューで、Excel、OneNote、またはPowerPointに埋め込みます。 コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。 詳細[については、「content Office アドイン](../design/content-add-ins.md)」を参照してください。

「 [webview」も参照してください](#webview)。

## <a name="content-delivery-network-cdn"></a>コンテンツ配信ネットワーク (CDN)

コンテンツ **配信ネットワークまたは****CDNは**、サーバーとデータ センターの分散ネットワークです。 通常、単一のサーバーまたはデータ センターと比較すると、リソースの可用性とパフォーマンスが向上します。

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (Contoso および Contoso University とも呼ばれる) は、Microsoft が会社とドメインの例として使用する架空の会社です。

## <a name="custom-function"></a>カスタム関数

カスタム **関数は**、カスタム アドインと一緒にパッケージ化されたユーザー Excel関数です。 カスタム関数を使用すると、開発者は、アドインの一部として JavaScript でそれらの関数を定義することで、Excel機能を超えて新しい関数を追加できます。 ユーザーは、Excel内のネイティブ関数と同じ方法でカスタム関数にExcel。 詳細については[、「Create custom functions in Excel](../excel/custom-functions-overview.md)」を参照してください。

## <a name="custom-functions-runtime"></a>カスタム関数ランタイム

カスタム **関数ランタイムは、** カスタム関数のみを実行する JavaScript ランタイムです。 UI は持たれ、API を操作Office.jsできません。 アドインにカスタム関数しか含めない場合、これは軽量で使いやすいランタイムです。 カスタム関数が作業ウィンドウまたは API を操作する必要がある場合は、Office.js JavaScript ランタイムを構成します。 詳細については、「[Office アドインを構成して共有 JavaScript ランタイムを使用する ](../develop/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

「 [JavaScript ランタイム、共有 JavaScript](#javascript-runtime) [ランタイム、共有ランタイム」も参照してください](#shared-javascript-runtime-shared-runtime)。

## <a name="host"></a>host

**ホスト** は通常、アプリケーションのOfficeします。 Office Office アドインをサポートする Office アプリケーション 、またはホストは、Excel、OneNote、Outlook、PowerPoint、Project、および Word です。

「application、[client](#application)[、Office](#client)[、Office」も参照してください](#office-application-office-client)。

## <a name="javascript-runtime"></a>JavaScript ランタイム

**JavaScript ランタイムは**、アドインが実行されるブラウザー ホスト環境です。 Mac Office Windows および Office では、JavaScript ランタイムは、Internet Explorer、Edge Legacy、Edge WebView2、Safari などの埋め込みブラウザー コントロール (または webview) です。 アドインの異なる部分は、個別の JavaScript ランタイムで実行されます。 たとえば、共有 JavaScript ランタイムを構成しない限り、アドイン コマンド、カスタム関数、作業ウィンドウ コードは通常、個別の JavaScript ランタイムを使用します。 詳細[については、「Officeアドイン](../concepts/browsers-used-by-office-web-add-ins.md)で使用されるブラウザー」を参照してください。

「カスタム関数ランタイム[、共有](#custom-functions-runtime) [JavaScript ランタイム、共有ランタイム、Webview」も](#shared-javascript-runtime-shared-runtime)[参照してください](#webview)。

## <a name="office-application-office-client"></a>Office アプリケーション、Office クライアント

**Officeクライアントは**、アプリケーションをOfficeします。 Office アドインをサポートする Office アプリケーション、またはクライアントは、Excel、OneNote、Outlook、PowerPoint、Project、および Word です。

「アプリケーション、[クライアント、](#application)[ホスト」も](#client)[参照してください](#host)。

## <a name="platform"></a>platform

プラットフォーム **は** 通常、アプリケーションを実行しているオペレーティング Officeします。 アドインをサポートOfficeプラットフォームには、Windows Mac、iPad、Web ブラウザーが含まれます。

## <a name="quick-start"></a>クイック スタート

クイック **スタートは** 、特定のプログラムの基本的な操作に必要な主要なスキルと知識の概要です。 アドインのOfficeでは、クイック スタートは、アプリケーションなどの特定のアプリケーション用のアドインの開発の概要Outlook。 クイック スタートには、アドイン開発者が約 5 分で完了できる一連の手順が含まれているので、アドインと機能の開発環境が機能します。

「チュートリアル」も [参照してください](#tutorial)。

## <a name="requirement-set"></a>要件セット

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>リボン、リボン ボタン

リボン **は** 、アプリケーションの機能をウィンドウの上部にある一連のタブまたはボタンに整理するコマンド バーです。 リボン **ボタンは** 、このシリーズ内のボタンの 1 つです。 詳細[については、「リボンを表示または非表示にする」Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions)を参照してください。

## <a name="runtime"></a>ランタイム

「 [JavaScript ランタイム」を参照してください](#javascript-runtime)。

## <a name="shared-javascript-runtime-shared-runtime"></a>共有 JavaScript ランタイム、共有ランタイム

共有 **JavaScript** ランタイム (共有 **ランタイム) を** 使用すると、作業ウィンドウ、アドイン コマンド、カスタム関数など、アドイン内のすべてのコードを同じ JavaScript ランタイムで実行し、作業ウィンドウが閉じても実行を続行できます。 詳細については、「Office アドインで共有 [JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md) ランタイムを使用する Office アドインを構成する」および「[Office ヒント](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/)」を参照してください。

「カスタム関数ランタイム[、](#custom-functions-runtime)[JavaScript ランタイム」も参照してください](#javascript-runtime)。

## <a name="task-pane"></a>作業ウィンドウ

作業ウィンドウは、通常、Excel、Outlook、PowerPoint、および Word 内のウィンドウの右側に表示されるインターフェイス サーフェス (webview) です。 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 ドキュメントに機能を直接埋め込む必要やできない場合は、作業ウィンドウを使用します。 詳細[については、「Officeアドインの作業ウィンドウ](../design/task-pane-add-ins.md)」を参照してください。

「 [webview」も参照してください](#webview)。

## <a name="tutorial"></a>チュートリアル

チュートリアル **は** 、ユーザーが製品または手順を使用する方法を学ぶのに役立つ教材です。 アドインOfficeコンテキストでは、チュートリアルでは、アドイン開発者が特定のアプリケーション (Excel など) の完全なアドイン開発プロセスをガイドします。 これには、20 以上の手順に従う必要があります。クイック スタートよりも時間の投資 [が大きくなります](#quick-start)。

「クイック スタート [」も参照してください](#quick-start)。

## <a name="ui-less-custom-function"></a>UI レスのカスタム関数

**UI レスのカスタム関数は、** カスタム関数ランタイムで実行されます。 UI は持たれ、API を操作Office.jsできません。

「カスタム関数[、カスタム関数ランタイム](#custom-function)[」も参照してください](#custom-functions-runtime)。

## <a name="web-add-in"></a>Web アドイン

**Web アドインは、** 既存のアドインのOffice用語です。 この用語は、Microsoft 365 ドキュメントでモダン Office アドインを VBA、COM、または VSTO などの他の種類のアドインと区別する必要がある場合に使用できます。

「アドイン」 [も参照してください](#add-in)。

## <a name="webview"></a>webview

**Webview は**、アプリケーション内の Web コンテンツを表示する要素またはビューです。 コンテンツ アドインと作業ウィンドウには、両方とも埋め込み Web ブラウザーが含まれているので、webviews の例として、Officeがあります。

「コンテンツ アドイン[、作業ウィンドウ」](#content-add-in)[も参照してください](#task-pane)。

## <a name="xll"></a>XLL

**XLL アドイン** は、ユーザー Excel機能を提供し、ファイル拡張子 **が .xll** のアドイン ファイルです。 XLL ファイルは、ダイナミック リンク ライブラリ (DLL) ファイルの一種で、このファイルを開くExcel。 XLL アドイン ファイルは、C または C++ で記述する必要があります。 カスタム関数は、XLL ユーザー定義関数に相当する最新の関数です。 カスタム関数は、プラットフォーム間でサポートを提供し、XLL ファイルと下位互換性があります。 詳細については [、「XLL ユーザー定義関数を使用してカスタム関数を拡張する](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) 」を参照してください。

「カスタム関数」 [も参照してください](#custom-function)。

## <a name="yeoman-generator-yo-office"></a>Yeoman ジェネレーター、yo office

アドイン[の Yeoman ジェネレーター Office](https://github.com/OfficeDev/generator-office)、オープンソースの [Yeoman](https://github.com/yeoman/yo) ツールを使用して、コマンド ラインをOfficeアドインを生成します。 `yo office`は、アドインの Yeoman ジェネレーターをOfficeコマンドです。このOfficeクイック スタートとチュートリアルでは、Yeoman ジェネレーターを使用します。

## <a name="see-also"></a>関連項目

- [Office アドインのその他のリソース](resources-links-help.md)