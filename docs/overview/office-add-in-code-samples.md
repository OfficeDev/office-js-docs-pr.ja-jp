---
title: Office アドインのコード サンプル
description: 独自のアドインの学習や作成に役立つ Office アドインのコード サンプルの一覧。
ms.date: 06/10/2022
localization_priority: high
ms.openlocfilehash: 16a1f92992c397772559468c27033aa58f6b6a6d
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423266"
---
# <a name="office-add-in-code-samples"></a>Office アドインのコード サンプル

これらのコード サンプルは、Office アドインを開発する場合のさまざまな機能の使用方法を学ぶサポートのために書かれています。

## <a name="getting-started"></a>はじめに

次のサンプルは、マニフェスト、HTML Web ページ、ロゴのみを使用して最も単純な Office アドインを構築する方法を示しています。 これらのコンポーネントは、Office アドインの基本的な部分です。 その他の開始情報については、[クイック スタート](../quickstarts/excel-quickstart-jquery.md)と[チュートリアル](/search/?terms=tutorial&scope=Office%20Add-ins)を参照してください。

- [Excel "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
- [Outlook "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
- [PowerPoint "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
- [Word "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

<br>

---

---

## <a name="blazor-webassembly"></a>Blazor WebAssembly

- [Blazor WebAssembly Excel アドインを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/excel-blazor-add-in)
- [Blazor WebAssembly Word アドインを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/word-blazor-add-in)

## <a name="excel"></a>Excel

| 名前                | 説明         |
|:--------------------|:--------------------|
| [Teams で開く](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-open-in-teams) | Microsoft Teams で、定義したデータを含む新しい Excel スプレッドシートを作成します。|
| [外部の Excel ファイルを挿入し、JSON データで設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-insert-file)  | 現在開いている Excel ブックに、外部の Excel ファイルの既存のテンプレートを挿入します。 次に、JSON Web サービスのデータをテンプレートに設定します。 |
| [リボンのカスタム コンテキスト タブを作成する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs) | Office UI のリボンでカスタム コンテクスト タブを作成します。 このサンプルでは、テーブルを作成し、ユーザーがテーブル内にフォーカスを移動させると、カスタム タブが表示されます。 ユーザーがテーブルの外に移動すると、カスタム タブは非表示になります。 |
| [Office アドイン アクション用のキーボード ショートカットを使用する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) | キーボード ショートカットを利用する基本的な Excel アドイン プロジェクトを設定します。 |
| [Web ワーカーを使用したカスタム関数のサンプル](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/web-worker) | カスタム関数で Web ワーカーを使用して、お使いの Office アドインの UI をブロックしないようにします。 |
| [ストレージ テクニックを使用してオフライン時に Office アドインからデータにアクセスする](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin) | ユーザー エクスペリエンスの接続が失われた場合に、お使いの Office アドイン向けに制限された機能を有効にする localStorage を実装します。 |
| [カスタム関数のバッチ処理パターン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)| 複数の呼び出しを単一の呼び出しにバッチ処理し、リモート サービスへのネットワーク呼び出しの回数を減らします。|

## <a name="outlook"></a>Outlook

| 名前                | 説明         |
|:--------------------|:--------------------|
| [添付ファイルを暗号化し、会議出席依頼の出席者を処理し、予定の日付/時刻の変更に対応](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | イベント ベースのアクティブ化を使用して、ユーザーが追加したときに添付ファイルを暗号化します。また、会議出席依頼で変更された受信者、および会議出席依頼の開始または終了の日時の変更にはイベント処理も使用します。 |
| [Outlook イベントベースのアクティブ化を使用して、外部受信者をタグ付けする](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | イベントベースのアクティブ化を使用して、ユーザーがメッセージ作成中に受信者を変更した場合に Outlook アドインを実行します。 このアドインでは、`appendOnSendAsync` API も使用して免責事項を追加します。 |
| [Outlook イベントベースのアクティブ化を使用して署名を設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | イベント ベースのアクティブ化を使用して、ユーザーが新しいメッセージまたは予定を作成するときに Outlook アドインを実行します。 アドインは、作業ウィンドウが開いていない場合でも、イベントに応答できます。 このアドインでは、`setSignatureAsync` API も使用します。 |
| [Outlook スマート アラートを使用する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories) | Outlook スマート アラートを使用して、新規メッセージまたは予定が送信される前に、必要な色の分類項目が適用されていることを確認します。 |

## <a name="word"></a>Word

| 名前                | 説明         |
|:--------------------|:--------------------|
| [Word アドインを使用して、Word 文書の OOXML コンテンツを取得、編集、および設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-get-set-edit-openxml) | このサンプルは、Word 文書の OOXML コンテンツを取得、編集、および設定する方法を示しています。 サンプル アドインは、独自のコンテンツ用に Office Open XML を取得し、独自に編集した Office Open XML スニペットをテストするため、スクラッチ パッドを提供します。|
| [Word アドインに Open XML を読み込んで書き込む](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)  | このサンプル アドインは、setSelectedDataAsync メソッドと ooxml 型変換を使用して、さまざまな種類のリッチ コンテンツ タイプを Word 文書に追加する方法を示します。 また、このアドインでは、サンプル コンテンツ タイプごとに Office Open XML マークアップをページ上に表示することもできます。 |

<br>

---

---

## <a name="authentication-authorization-and-single-sign-on-sso"></a>認証、承認、シングル サインオン (SSO)

| 名前                | 説明         |
|:--------------------|:--------------------|
| [シングル サインオン (SSO) サンプル Outlook アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) | Office の SSO 機能を使用して、アドインが Microsoft Graph データにアクセスできるようにします。|
| [Office アドインの Microsoft Graph と msal.js を使用して OneDrive データを取得する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | バックエンドのないシングル ページ アプリケーション (SPA) として、Microsoft Graph に接続する Office アドインを作成し、OneDrive for Business に保存されているブックにアクセスして、スプレッドシートを更新します。  |
| [Microsoft Graph への Office アドイン認証](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Microsoft Graph に接続して OneDrive for Business に保存されているブックにアクセスし、スプレッドシートを更新する Microsoft Office アドインの作成方法について学習します。 |
| [Microsoft Graph への Outlook アドイン認証](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)。 | Microsoft Graph に接続して OneDrive for Business に保存されているブックにアクセスし、新しいメール メッセージを作成する Outlook アドインを作成します。 |
| [ASP.NET を使用したシングル サインオン (SSO) Office アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | Office.js で `getAccessToken` API を使用して、アドインが Microsoft Graph データにアクセスできるようにします。このサンプルは ASP.NET で作成されています。 |
| [Node.js を使用したシングル サインオン (SSO) Office アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Office.js で `getAccessToken` API を使用して、アドインが Microsoft Graph データにアクセスできるようにします。このサンプルは Node.js で作成されています。|

## <a name="shared-runtime"></a>共有ランタイム

| 名前                | 説明         |
|:--------------------|:--------------------|
| [グローバル データを共有ランタイムと共有する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-global-state) | 共有ランタイムを使用して、リボン ボタン、作業ウィンドウ、カスタム関数のコードを単一のブラウザー ランタイムで実行する基本的なプロジェクトを設定します。 |
| [リボンと作業ウィンドウ UI を管理し、開いたドキュメントでコードを実行する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario) | アドインの状態に基づいて有効になる状況依存のリボンのボタンを作成します。 |

<br>

---

---

## <a name="additional-samples"></a>追加サンプル

| 名前                | 説明         |
|:--------------------|:--------------------|
| [共有ライブラリを使用して、Visual Studio Tools for Office アドインを Office Web アドインに移行する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) | VSTO アドインから Office アドインに移行する場合に、コードを再利用するための戦略を提供します。 |
| [Azure 関数を Excel カスタム関数と統合する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AzureFunction) | Azure 関数とカスタム関数を統合して、クラウドに移行したり追加サービスを統合したりします。 |
| [動的 DPI コードのサンプル](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/dynamic-dpi) | COM、VSTO、Office アドインの DPI 変更を処理するためのサンプルのコレクション。 |

## <a name="next-steps"></a>次の手順

Microsoft 365 開発者プログラムに参加します。Microsoft 365 プラットフォームのソリューションを構築するために必要な無料のサンドボックス、ツール、およびその他のリソースを入手してください。

- [無料の開発者サンドボックス](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) 90 日間の更新可能な無料の Microsoft 365 E5 開発者サブスクリプションを取得します。
- [サンプル データ パック](https://developer.microsoft.com/microsoft-365/dev-program#Sample) ソリューションの構築に役立つユーザー データとコンテンツをインストールして、サンドボックスを自動的に構成します。
- [専門家への相談](https://developer.microsoft.com/microsoft-365/dev-program#Experts) コミュニティ イベントにアクセスして、Microsoft 365 エキスパートから学びます。
- [個人用の推奨事項](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) 個人用ダッシュボードから開発者のリソースをすばやく見つけます。
