---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、Office アドインの API アクセスレベルを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 603494b61ef126b35cb5cdff8c5f5b911bd25840
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611492"
---
# <a name="permissions-element"></a>Permissions 要素

Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

コンテンツ アドインおよび作業ウィンドウ アドインの場合:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

詳細については、「[アドインで API を使用するためのアクセス許可を要求](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)する」と「 [Outlook アドインのアクセス許可につい](../../outlook/understanding-outlook-add-in-permissions.md)て」を参照してください。
