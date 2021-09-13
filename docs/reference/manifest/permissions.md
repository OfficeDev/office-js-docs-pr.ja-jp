---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、アドインの API アクセス レベルOffice指定します。
ms.date: 06/26/2020
ms.localizationpriority: medium
ms.openlocfilehash: a472d7a6f375c3a171fdd529b993aaf2c6109ce9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154852"
---
# <a name="permissions-element"></a>Permissions 要素

Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

コンテンツ アドインおよび作業ウィンドウ アドインの場合:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

詳細については、「コンテンツアドインと作業ウィンドウ アドインでの[API](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)使用のアクセス許可の要求」および「Outlookについて」を[参照してください](../../outlook/understanding-outlook-add-in-permissions.md)。
