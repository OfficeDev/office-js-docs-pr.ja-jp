---
title: SSO を使用する Office アドインをMicrosoft ID プラットフォームに登録する
description: Office アドインを Microsoft ID プラットフォーム に登録して、Word、Excel、PowerPoint、Outlook で SSO を使用する方法について説明します。
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0aab7d421ac57d1436d68c659f5d820717bcb846
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842097"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>シングル サインオン (SSO) を使用する Office アドインを Microsoft ID プラットフォームに登録する

この記事では、SSO を使用できるように Office アドインをMicrosoft ID プラットフォームに登録する方法について説明します。 アドインの開発を開始するときにアドインを登録して、テストまたは運用に進むときに、既存の登録を変更したり、アドインの開発、テスト、および運用バージョン用に個別の登録を作成したりできます。

次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。

|情報  |例  |プレースホルダー  |
|---------|---------|---------|
|人間が判読できるアドインの名前です  (一意であることが推奨されますが、必須ではありません)。|`Contoso Marketing Excel Add-in (Prod)`|該当なし|
|登録プロセスの一環として Azure によって生成されるアプリケーション ID。|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|アドインの完全修飾ドメイン名 (プロトコルを除く) です。 *所有しているドメインを使用する必要があります。* この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。 ドメインは、アドインのマニフェストのセクションの URL **\<Resources\>** で使用されているように、サブドメインも含めて同じである必要があります。|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|アドインに必要なMicrosoft ID プラットフォームと Microsoft Graph に対するアクセス許可。 (`profile` は常に必須です)。|`profile`, `Files.Read.All`|N/A|

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]