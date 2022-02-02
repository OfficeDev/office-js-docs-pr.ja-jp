---
title: アドインに管理者の同意を許可する
description: 管理者の同意をアドインに付与する方法について学習する
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2c3a82db390ed28c1eb8194a78f2c9fa787aeede
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320131"
---
# <a name="grant-administrator-consent-to-the-add-in"></a>アドインに管理者の同意を許可する

> [!NOTE]
> この手順が必要とされるのは、アドインを開発しているときだけです。 実稼働アドインを AppSource または Microsoft 365 管理センター に展開すると、ユーザーは個別に信頼するか、管理者がインストール時に組織に同意します。

アドインを登録 *した後* で [、この手順を実行します](../develop/register-sso-add-in-aad-v2.md)。

1. [Azure portal - アプリ登録ページを参照して](https://go.microsoft.com/fwlink/?linkid=2083908)、アプリの登録を表示します。

1. 管理者資格情報を ***使用してテナント*** にサインインMicrosoft 365します。 たとえば、MyName@contoso.onmicrosoft.com です。

1. 表示名が表示されているアプリを選択 **$ADD-IN-NAME$を指定します**。

1. [**$ADD-IN-NAME$**] ページで **、[API** のアクセス許可] を選択し、[構成されたアクセス許可] セクションで、[[テナント名] に対する管理者の同意を許可する] **を選択します**。 表示 **される確認に** 対して [はい] を選択します。

> [!NOTE]
> 開発者アカウントを使用している場合は、この手順をベスト [プラクティスMicrosoft 365勧めします](https://developer.microsoft.com/microsoft-365/dev-program)。 ただし、必要に応じて、開発中の SSO アドインをサイドロードし、ユーザーに同意フォームを求めるメッセージを表示することもできます。 詳細については、「[Sideload on Windows」](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)と「サイドロード オン Office on the web」 を[参照してください](../testing/sideload-office-add-ins-for-testing.md)。
