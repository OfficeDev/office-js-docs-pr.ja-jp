---
title: SSO をOfficeするアドインをそのアドインに登録Microsoft ID プラットフォーム
description: Excel Word、Office、PowerPoint、および Outlook で SSO を使用する Microsoft ID プラットフォーム アドインを登録する方法についてOutlook。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: b11ce5130e020b049038631b9ae1c0e62fdadeab
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320248"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>シングル サインオンOffice (SSO) を使用するアドインをユーザーに登録Microsoft ID プラットフォーム

この記事では、SSO を使用Officeアドインを Microsoft ID プラットフォームする方法について説明します。 開発を開始する際にアドインを登録し、テストまたは実稼働に進むときに、既存の登録を変更したり、アドインの開発、テスト、および実稼働バージョンの個別の登録を作成したりすることができます。

次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。

|情報  |例  |プレースホルダー  |
|---------|---------|---------|
|人間が判読できるアドインの名前です  (一意であることが推奨されますが、必須ではありません)。|`Contoso Marketing Excel Add-in (Prod)`|該当なし|
|Azure が登録プロセスの一環として生成するアプリケーション ID。|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|アドインの完全修飾ドメイン名 (プロトコルを除く) です。 *所有しているドメインを使用する必要があります。* この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。 このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|アドインが必要とするMicrosoft ID プラットフォーム Microsoft Graphアクセス許可。 (`profile` は常に必須です)。|`profile`, `Files.Read.All`|N/A|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
