---
title: マニフェスト ファイルの Scopes 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771787"
---
# <a name="scopes-element"></a><span data-ttu-id="be15c-102">Scopes 要素</span><span class="sxs-lookup"><span data-stu-id="be15c-102">Scopes element</span></span>

<span data-ttu-id="be15c-103">アドインで必要な Microsoft Graph に対するアクセス許可が含まれます。</span><span class="sxs-lookup"><span data-stu-id="be15c-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="be15c-104">AppSource は、スコープ要素を使用して同意ダイアログボックスを作成します。</span><span class="sxs-lookup"><span data-stu-id="be15c-104">AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="be15c-105">ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。</span><span class="sxs-lookup"><span data-stu-id="be15c-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="be15c-106">子要素</span><span class="sxs-lookup"><span data-stu-id="be15c-106">Child elements</span></span>

|  <span data-ttu-id="be15c-107">要素</span><span class="sxs-lookup"><span data-stu-id="be15c-107">Element</span></span> |  <span data-ttu-id="be15c-108">型</span><span class="sxs-lookup"><span data-stu-id="be15c-108">Type</span></span>  |  <span data-ttu-id="be15c-109">説明</span><span class="sxs-lookup"><span data-stu-id="be15c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="be15c-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="be15c-110">**Scope**</span></span>                |  <span data-ttu-id="be15c-111">string</span><span class="sxs-lookup"><span data-stu-id="be15c-111">string</span></span>     |   <span data-ttu-id="be15c-112">Microsoft Graph に対するアクセス許可の名前。たとえば、Files.Read.All です。</span><span class="sxs-lookup"><span data-stu-id="be15c-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="be15c-113">例</span><span class="sxs-lookup"><span data-stu-id="be15c-113">Example</span></span>

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
