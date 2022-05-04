---
title: Office アドインを管理する
description: 互換性に対する取り組みと、アドインを最新の状態に保つ方法について説明します。
ms.date: 04/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 55da05d5c0b220adbeb0b4dbe248aa79f05b6b74
ms.sourcegitcommit: 5bf28c447c5b60e2cc7e7a2155db66cd9fe2ab6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/04/2022
ms.locfileid: "65187358"
---
# <a name="maintain-your-office-add-in"></a>Office アドインを管理する

アドインを発行した後は、アップストリーム ライブラリからの重要な変更を最新の状態に保つ必要があります。 セキュリティの問題にパッチを適用することは、お客様の信頼を築くために不可欠です。 これらの変更は公開されたマニフェストには影響しないため、お客様はアドインの最新バージョンを取得するためのアクションを実行する必要はありません。

## <a name="breaking-changes-in-officejs"></a>Office.jsの重大な変更

Microsoft 365開発者プラットフォームは、アドインの互換性を確保することに取り組んでいます。 API の表面と動作に重大な変更を加えないように努めています。 ただし、セキュリティや信頼性を確保するために、重大な更新を行う必要がある場合があります。 このようなまれなケースでは、アドインのユーザーが影響を受けないようにするために、次の手順が実行されます。

- 影響を受ける機能と推奨される変更について説明したお知らせは、[Microsoft 365開発者ブログ](https://devblogs.microsoft.com/microsoft365dev/)で行われます。
- [アドインが AppSource](/office/dev/store/submit-to-appsource-via-partner-center) で公開されている場合は、指定した情報を通じて連絡を受けます。
- 可能であれば、影響を受けるMicrosoft 365 テナントの管理者 ([開発者テナントを](https://developer.microsoft.com/microsoft-365/dev-program)含む) は[、メッセージ センター](/microsoft-365/admin/manage/message-center)から連絡を受けます。 AppSource の外部で公開されたアドイン ソリューションのプロバイダーに問い合わせるのは、管理者の責任です。

### <a name="app-assure"></a>App Assure

Microsoft [の App Assure](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) サービスは、アプリケーションの互換性という Microsoft の約束を果たします。アプリはWindowsとMicrosoft 365 Appsで動作します。 App Assure エンジニアは、追加コストなしで発生する可能性のある問題の解決に役立ちます。

アプリの互換性の問題が発生した場合は、App Assure エンジニアが協力して問題の解決に役立ちます。 エキスパートは次の情報を提供します。

- 根本原因のトラブルシューティングと特定に役立ちます。
- アプリケーションの互換性の問題を修復するのに役立つガイダンスを提供します。
- アプリの一部を修復するために、お客様に代わって独立系ソフトウェア ベンダー (ISV) と連携し、最新バージョンの製品で機能するようにします。
- Microsoft 製品エンジニアリング チームと連携して、製品のバグを修正します。

App Assure の詳細については、「App Assure で[アプリをMicrosoft Edgeする: ヒントとテクニック](https://techcommunity.microsoft.com/t5/video-hub/bring-your-apps-to-microsoft-edge-with-app-assure-tips-and/ba-p/2167619)」を参照してください。 App Assure との互換性に関する要求を送信するには、[Microsoft FastTrack登録フォーム](https://aka.ms/AppAssureRequest)に入力するか[、achelp@microsoft.com](mailto:achelp@microsoft.com) に電子メールを送信します。

## <a name="changes-to-yeoman-templates-and-web-dependencies"></a>Yeoman テンプレートと Web 依存関係に対する変更

[Office アドイン用 Yeoman Generator は、](../develop/yeoman-generator-overview.md)Microsoft や他のライブラリの数に依存しています。 これらのライブラリは、Microsoft 365アクティビティとは別に更新されます。 ジェネレーターを使用して作成されたすべてのプロジェクトは、アドインの開発、発行、および保守を行う際に最新の状態に保たれる必要があります。 次のツールは、プロジェクトが依存ライブラリのセキュリティで保護されたバージョンを使用していることを確認するのに役立ちます。

- [npm 監査](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot とその他のGitHubセキュリティ機能](https://github.com/features/security)

このガイダンスは、[Office アドイン のコード サンプル](https://github.com/OfficeDev/Office-Add-in-samples)やその他のソースから取得したサンプルのコピーにも適用されます。

### <a name="officejs-npm-package"></a>office.js NPM パッケージ

[office-js NPM パッケージ](https://www.npmjs.com/package/@microsoft/office-js)は、[Office.js コンテンツ配信ネットワーク (CDN)](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library) でホストされているもののコピーです。 これは、CDNへの直接アクセスが不可能なシナリオを対象としています。 NPM パッケージは、office.jsへのバージョン管理された参照を提供することを目的としていません。 Office JavaScript API の最新バージョンを確実に使用するには、常にCDNを使用することを強くお勧めします。

## <a name="current-best-practices"></a>現在のベスト プラクティス

下位互換性の維持に努めていますが、パターンとプラクティスは継続的に進化することをお勧めします。 ドキュメントは、現在のベスト プラクティスを示すために努めています。 既存の機能を向上させる可能性のある新機能について常に情報を得るには、毎月[のOffice アドインCommunity通話に](../overview/office-add-ins-community-call.md)参加してください。
