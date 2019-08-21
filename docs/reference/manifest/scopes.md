---
title: マニフェスト ファイルの Scopes 要素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1e36bdcd0cdcaa8c842e924c2543d56bdc4e26a7
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477734"
---
# <a name="scopes-element"></a><span data-ttu-id="7f0c1-102">Scopes 要素</span><span class="sxs-lookup"><span data-stu-id="7f0c1-102">Scopes element</span></span>

<span data-ttu-id="7f0c1-103">アドインが外部リソース (Microsoft Graph など) に対して必要とするアクセス許可が含まれます。</span><span class="sxs-lookup"><span data-stu-id="7f0c1-103">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="7f0c1-104">Microsoft Graph がリソースの場合、AppSource はスコープ要素を使用して同意ダイアログボックスを作成します。</span><span class="sxs-lookup"><span data-stu-id="7f0c1-104">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="7f0c1-105">ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。</span><span class="sxs-lookup"><span data-stu-id="7f0c1-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="7f0c1-106">**スコープ**は、マニフェスト内の[Webapplicationinfo](webapplicationinfo.md)要素と[Authorization](authorization.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="7f0c1-106">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7f0c1-107">子要素</span><span class="sxs-lookup"><span data-stu-id="7f0c1-107">Child elements</span></span>

|  <span data-ttu-id="7f0c1-108">要素</span><span class="sxs-lookup"><span data-stu-id="7f0c1-108">Element</span></span> |  <span data-ttu-id="7f0c1-109">必須</span><span class="sxs-lookup"><span data-stu-id="7f0c1-109">Required</span></span>  |  <span data-ttu-id="7f0c1-110">説明</span><span class="sxs-lookup"><span data-stu-id="7f0c1-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7f0c1-111">**Scope**</span><span class="sxs-lookup"><span data-stu-id="7f0c1-111">**Scope**</span></span>                |  <span data-ttu-id="7f0c1-112">はい</span><span class="sxs-lookup"><span data-stu-id="7f0c1-112">Yes</span></span>     |   <span data-ttu-id="7f0c1-113">アクセス許可の名前。たとえば、[すべて] または [プロファイル] を参照します。</span><span class="sxs-lookup"><span data-stu-id="7f0c1-113">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="7f0c1-114">例</span><span class="sxs-lookup"><span data-stu-id="7f0c1-114">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
