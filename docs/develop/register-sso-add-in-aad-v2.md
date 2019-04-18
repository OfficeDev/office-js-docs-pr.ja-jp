---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: ''
ms.date: 04/10/2019
localization_priority: Priority
ms.openlocfilehash: a98fb7e9f073024804f577057fde83d1bdc83273
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914250"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する

この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。 開発を開始する前に、アドインを登録する必要があります。 テストまたは運用環境に進んだ場合、既存の登録を変更するか、アドインの開発、テスト、および運用バージョン用に別の登録を作成できます。

次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。

|情報  |例  |プレースホルダー  |
|---------|---------|---------|
|人間が判読できるアドインの名前です  (一意であることが推奨されますが、必須ではありません)。|`Contoso Marketing Excel Add-in (Prod)`|**$ADD-IN-NAME$**|
|アドインの完全修飾ドメイン名 (プロトコルを除く) です。 *所有しているドメインを使用する必要があります。* この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。 このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。|`localhost:6789`, `addins.contoso.com`|**$FQDN-WITHOUT-PROTOCOL$**|
|ご使用のアドインに必要な AAD および Microsoft Graph へのアクセス許可です  (`profile` は常に必須です)。|`profile`, `Files.Read.All`|N/A|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
