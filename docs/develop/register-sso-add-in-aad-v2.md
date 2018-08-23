---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437256"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する

この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。 アドインの開発を開始するときには、アドインを登録する必要があります。 テストまたは運用に進むと、アドインの開発、テスト、および運用バージョン用に、既存の登録を変更するか、別々の登録を作成するかできます。 

次の表には、このプロシージャを実行するために必要な情報と、指示に表示される対応するプレースホルダーを挙げてあります。 

|情報  |例  |プレースホルダー  |
|---------|---------|---------|
|アドインの読みやすい名前。 (一意であることが推奨されます。ただし、一意でなくてもかまいません。)    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|アドインの完全修飾ドメイン名（プロトコルを除く）。 *自分が所有するドメインを使用する必要があります。* このため、 `azurewebsites.net` や `cloudapp.net` のような特定のよく知られたドメインは使用できません。   |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|アドインに必要な AAD と Microsoft Graph へのアクセス許可。 （`profile` が常に必要です。）    |`profile`, `Files.Read.All`         |該当なし         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]