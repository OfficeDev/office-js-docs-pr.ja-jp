---
title: アドインがインストールされたときに作業ウィンドウを自動的に開く
description: インストール時に自動的に開く Office アドインを構成する方法について説明します。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d6ff4b8b5b68236d435ec91b2dcbe121f211081d
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674768"
---
# <a name="automatically-open-a-task-pane-when-an-add-in-is-installed"></a>アドインがインストールされたときに作業ウィンドウを自動的に開く

アドインの作業ウィンドウは、インストール直後に起動するように構成できます。 この機能により、使用量が増加します。 

既定では、アドイン コマンドを含 *まない* 作業ウィンドウ [アドイン](../design/add-in-commands.md) は、インストール直後に作業ウィンドウを開きます。 ただし、アドインに 1 つ以上のアドイン コマンドがある場合、ユーザーには新しいアドインが通知されますが、アドインは自動的に起動しません。 この過去の既定の動作は変化しているため、アドイン コマンドを持つアドインは、状況によっては自動的に起動します。 さらに、アドインに複数の作業ウィンドウ ページがある場合は、インストール時にアドインを起動するかどうか、およびインストールされている場合は作業ウィンドウで開くページを制御できます。

> [!NOTE]
> 
> - この機能は現在、Office on the webでのみ使用できます。 この動作を他のプラットフォームに持ち込むことに取り組んでいますが、現在でも、以前に説明した従来の既定の動作が示されています。
> - この機能は、エンド ユーザーがインストールしたアドインにのみ適用され、一元的に展開されたアドインには適用されません。
> - この機能は、コンテンツ アドインまたはメール (Outlook) アドインには適用されません。
> - この機能は、" [作業ウィンドウ コマンド" 型](../design/add-in-commands.md#types-of-add-in-commands)のアドイン コマンドが少なくとも 1 つ存在するアドインにのみ適用されます。

## <a name="new-behavior"></a>新しい動作

新しい動作は次のとおりです。

- アドインに 1 つの [作業ウィンドウ コマンド](../design/add-in-commands.md#types-of-add-in-commands)しかない場合は、アドインのリボン タブが選択され、インストール時に作業ウィンドウが自動的に開きます。 何も構成する必要はありません。
- アドインに複数の作業ウィンドウ コマンドがあり、1 つが既定の作業ウィンドウとして構成されている場合 ( [既定の作業ウィンドウの構成](#configure-default-task-pane)を参照)、アドインのリボン タブが選択され、インストール時に既定の作業ウィンドウが自動的に開きます。
- アドインに複数の作業ウィンドウ コマンドが含まれているが、既定として構成されていない場合は、インストール時にアドインのリボン タブが自動的に選択され、新しいアドインのユーザーに通知する吹き出しが表示されますが、作業ウィンドウは開かなくなります。 これは、履歴の既定の動作と同じです。

> [!NOTE]
> 何らかの理由で、作業ウィンドウを起動するアドイン コマンドを、起動時にユーザーが手動で選択できない場合 (起動時 [に無効にするように構成されている](../design/disable-add-in-commands.md) 場合など) は、構成に関係なく自動的に開かれることはありません。 

## <a name="configure-default-task-pane"></a>既定の作業ウィンドウを構成する

既定として作業ウィンドウを指定するには、 [要素の最初の子として TaskpaneId](/javascript/api/manifest/action#taskpaneid) 要素を **\<Action\>** 追加し、その値を **Office.AutoShowTaskpaneWithDocument** に設定します。 次に例を示します。

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

> [!TIP]
> ユーザーがドキュメントを再度開くたびにアドインを自動的に起動する場合は、さらに構成手順を実行する必要があります。 この機能を使用するタイミングの詳細とアドバイスについては、「 [ドキュメントで作業ウィンドウを自動的に開く](automatically-open-a-task-pane-with-a-document.md)」を参照してください。 

## <a name="see-also"></a>関連項目

- [ドキュメントで作業ウィンドウを自動的に開く](automatically-open-a-task-pane-with-a-document.md)
