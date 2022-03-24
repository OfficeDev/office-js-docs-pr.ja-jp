---
title: SSO をOfficeするアドインをアプリに登録Microsoft ID プラットフォーム
description: Word、Office、PowerPoint Microsoft ID プラットフォーム、および Outlook で SSO を使用する Excel アドインを Outlook。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e408a57534437f0d0fe0c5fb3b4ab844f7dde9ac
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743382"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>シングル サインオンOffice (SSO) を使用するアドインを、そのアドインに登録Microsoft ID プラットフォーム

この記事では、SSO を使用Officeアドインを Microsoft ID プラットフォームする方法について説明します。 開発を開始する際にアドインを登録し、テストまたは実稼働に進むときに、既存の登録を変更したり、アドインの開発、テスト、および実稼働バージョンの個別の登録を作成したりすることができます。

次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。

|情報  |例  |プレースホルダー  |
|---------|---------|---------|
|人間が判読できるアドインの名前です  (一意であることが推奨されますが、必須ではありません)。|`Contoso Marketing Excel Add-in (Prod)`|該当なし|
|Azure が登録プロセスの一環として生成するアプリケーション ID。|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|アドインの完全修飾ドメイン名 (プロトコルを除く) です。 *所有しているドメインを使用する必要があります。* この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。 このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|アドインが必要とするMicrosoft ID プラットフォーム Microsoft Graphアクセス許可。 (`profile` は常に必須です)。|`profile`, `Files.Read.All`|N/A|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
