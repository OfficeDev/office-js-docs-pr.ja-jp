---
title: Office アドインのコード サンプル
description: 独自のアドインの学習や作成に役立つ Office アドインのコード サンプルの一覧。
ms.date: 11/18/2021
localization_priority: high
ms.openlocfilehash: de86cacef6ae991af219e25351d32b8001ae6d39
ms.sourcegitcommit: 63e9ef5e210e353700a3d27cfbd5ed838ce55d01
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/24/2022
ms.locfileid: "62177716"
---
# <a name="office-add-in-code-samples"></a>Office アドインのコード サンプル

これらのコード サンプルは、Office アドインを開発する場合のさまざまな機能の使用方法を学ぶサポートのために書かれています。

## <a name="getting-started"></a>はじめに

次のサンプルは、マニフェスト、HTML Web ページ、ロゴのみを使用して最も単純な Office アドインを構築する方法を示しています。 これらのコンポーネントは、Office アドインの基本的な部分です。 その他の開始情報については、[クイック スタート](../quickstarts/excel-quickstart-jquery.md)と[チュートリアル](/search/?terms=tutorial&scope=Office%20Add-ins)を参照してください。

* [Excel "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
* [Outlook "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
* [PowerPoint "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
* [Word "Hello world" アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

## <a name="outlook"></a>Outlook

| 名前                | 説明         |
|:--------------------|:--------------------|
| [添付ファイルを暗号化し、会議出席依頼の出席者を処理し、予定の日付/時刻の変更に対応](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | イベント ベースのアクティブ化を使用して、ユーザーが追加したときに添付ファイルを暗号化します。 また、会議出席依頼で変更された受信者、および会議出席依頼の開始または終了の日時の変更にはイベント処理も使用します。 |
| [Outlook イベントベースのアクティブ化を使用して外部受信者をタグ付けする (プレビュー)](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | イベントベースのアクティブ化を使用して、ユーザーがメッセージ作成中に受信者を変更した場合に Outlook アドインを実行します。 このアドインでは、`appendOnSendAsync` API も使用して免責事項を追加します。 |
| [Outlook イベントベースのアクティブ化を使用して署名を設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | イベント ベースのアクティブ化を使用して、ユーザーが新しいメッセージまたは予定を作成するときに Outlook アドインを実行します。 アドインは、作業ウィンドウが開いていない場合でも、イベントに応答できます。 このアドインでは、`setSignatureAsync` API も使用します。 |

## <a name="excel"></a>Excel

| 名前                | 説明         |
|:--------------------|:--------------------|
| [Teams で開く](/samples/officedev/pnp-officeaddins/office-excel-add-in-open-in-teams/) | Microsoft Teams で、定義したデータを含む新しい Excel スプレッドシートを作成します。|
| [外部の Excel ファイルを挿入し、JSON データで設定する](/samples/officedev/pnp-officeaddins/excel-add-in-insert-external-file/)  | 現在開いている Excel ブックに、外部の Excel ファイルの既存のテンプレートを挿入します。 次に、JSON Web サービスのデータをテンプレートに設定します。 |
| [リボンのカスタム コンテキスト タブを作成する](/samples/officedev/pnp-officeaddins/office-add-in-contextual-tabs/) | Office UI のリボンでカスタム コンテクスト タブを作成します。 このサンプルでは、テーブルを作成し、ユーザーがテーブル内にフォーカスを移動させると、カスタム タブが表示されます。 ユーザーがテーブルの外に移動すると、カスタム タブは非表示になります。 |
| [Office アドイン アクション用のキーボード ショートカットを使用する](/samples/officedev/pnp-officeaddins/office-add-in-keyboard-shortcuts) | キーボード ショートカットを利用する基本的な Excel アドイン プロジェクトを設定します。 |
| [Web ワーカーを使用したカスタム関数のサンプル](/samples/officedev/pnp-officeaddins/excel-custom-function-web-worker-pattern/) | カスタム関数で Web ワーカーを使用して、お使いの Office アドインの UI をブロックしないようにします。 |
| [ストレージ テクニックを使用してオフライン時に Office アドインからデータにアクセスする](/samples/officedev/pnp-officeaddins/use-storage-techniques-to-access-data-from-an-office-add-in-when-offline/) | ユーザー エクスペリエンスの接続が失われた場合に、お使いの Office アドイン向けに制限された機能を有効にする localStorage を実装します。 |
| [カスタム関数のバッチ処理パターン](/samples/officedev/pnp-officeaddins/excel-custom-function-batching-pattern/)| 複数の呼び出しを単一の呼び出しにバッチ処理し、リモート サービスへのネットワーク呼び出しの回数を減らします。|

## <a name="shared-javascript-runtime"></a>共有 JavaScript ランタイム

| 名前                | 説明         |
|:--------------------|:--------------------|
[グローバル データを共有ランタイムと共有する](/samples/officedev/pnp-officeaddins/office-add-in-shared-runtime-global-data/) | 共有ランタイムを使用して、リボン ボタン、作業ウィンドウ、カスタム関数のコードを単一のブラウザー ランタイムで実行する基本的なプロジェクトを設定します。 |
| [リボンと作業ウィンドウ UI を管理し、開いたドキュメントでコードを実行する](/samples/officedev/pnp-officeaddins/office-add-in-ribbon-task-pane-ui/) | アドインの状態に基づいて有効になる状況依存のリボンのボタンを作成します。 |

## <a name="authentication-authorization-and-single-sign-on-sso"></a>認証、承認、シングル サインオン (SSO)

| 名前                | 説明         |
|:--------------------|:--------------------|
| [シングル サインオン (SSO) サンプル Outlook アドイン](/samples/officedev/pnp-officeaddins/outlook-add-in-sso-aspnet/) | Office の SSO 機能を使用して、アドインが Microsoft Graph データにアクセスできるようにします。|
| [Office アドインの Microsoft Graph と msal.js を使用して OneDrive データを取得する](/samples/officedev/pnp-officeaddins/office-add-in-auth-graph-react/) | バックエンドのないシングル ページ アプリケーション (SPA) として、Microsoft Graph に接続する Office アドインを作成し、OneDrive for Business に保存されているブックにアクセスして、スプレッドシートを更新します。  |
| [Microsoft Graph への Office アドイン認証](/samples/officedev/pnp-officeaddins/office-add-in-auth-aspnet-graph/) | Microsoft Graph に接続して OneDrive for Business に保存されているブックにアクセスし、スプレッドシートを更新する Microsoft Office アドインの作成方法について学習します。 |
| [Microsoft Graph への Outlook アドイン認証](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)。 | Microsoft Graph に接続して OneDrive for Business に保存されているブックにアクセスし、新しいメール メッセージを作成する Outlook アドインを作成します。 |
| [ASP.NET を使用したシングル サインオン (SSO) Office アドイン](/samples/officedev/pnp-officeaddins/office-add-in-sso-aspnet/) | Office.js で `getAccessToken` API を使用して、アドインが Microsoft Graph データにアクセスできるようにします。このサンプルは ASP.NET で作成されています。 |
| [Node.js を使用したシングル サインオン (SSO) Office アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Office.js で `getAccessToken` API を使用して、アドインが Microsoft Graph データにアクセスできるようにします。このサンプルは Node.js で作成されています。|

## <a name="additional-samples"></a>追加サンプル

| 名前                | 説明         |
|:--------------------|:--------------------|
|[共有ライブラリを使用して、Visual Studio Tools for Office アドインを Office Web アドインに移行する](/samples/officedev/pnp-officeaddins/vsto-shared-library-excel/) |VSTO アドインから Office アドインに移行する場合に、コードを再利用するための戦略を提供します。 |
| [Azure 関数を Excel カスタム関数と統合する](/samples/officedev/pnp-officeaddins/azure-function-with-excel-custom-function/) | Azure 関数とカスタム関数を統合して、クラウドに移行したり追加サービスを統合したりします。 |
|[動的 DPI コードのサンプル](/samples/officedev/pnp-officeaddins/dynamic-dpi-code-samples/) |COM、VSTO、Office アドインの DPI 変更を処理するためのサンプルのコレクション。 |

## <a name="next-steps"></a>次の手順

Microsoft 365 開発者プログラムに参加します。Microsoft 365 プラットフォームのソリューションを構築するために必要な無料のサンドボックス、ツール、およびその他のリソースを入手してください。

- [無料の開発者サンドボックス](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) 90 日間の更新可能な無料の Microsoft 365 E5 開発者サブスクリプションを取得します。
- [サンプル データ パック](https://developer.microsoft.com/microsoft-365/dev-program#Sample) ソリューションの構築に役立つユーザー データとコンテンツをインストールして、サンドボックスを自動的に構成します。
- [専門家への相談](https://developer.microsoft.com/microsoft-365/dev-program#Experts) コミュニティ イベントにアクセスして、Microsoft 365 エキスパートから学びます。
- [個人用の推奨事項](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) 個人用ダッシュボードから開発者のリソースをすばやく見つけます。
