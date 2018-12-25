---
title: マニフェスト ファイルの Scopes 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432642"
---
# <a name="scopes-element"></a><span data-ttu-id="09674-102">Scopes 要素</span><span class="sxs-lookup"><span data-stu-id="09674-102">Scopes element</span></span>

<span data-ttu-id="09674-103">アドインで必要な Microsoft Graph に対するアクセス許可が含まれます。</span><span class="sxs-lookup"><span data-stu-id="09674-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="09674-104">Office ストアは、Scopes 要素を使用して同意ダイアログ ボックスを作成します。</span><span class="sxs-lookup"><span data-stu-id="09674-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="09674-105">ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。</span><span class="sxs-lookup"><span data-stu-id="09674-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="09674-106">子要素</span><span class="sxs-lookup"><span data-stu-id="09674-106">Child elements</span></span>

|  <span data-ttu-id="09674-107">要素</span><span class="sxs-lookup"><span data-stu-id="09674-107">Element</span></span> |  <span data-ttu-id="09674-108">支払期日</span><span class="sxs-lookup"><span data-stu-id="09674-108">Type</span></span>  |  <span data-ttu-id="09674-109">説明</span><span class="sxs-lookup"><span data-stu-id="09674-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="09674-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="09674-110">**Scope**</span></span>                |  <span data-ttu-id="09674-111">string</span><span class="sxs-lookup"><span data-stu-id="09674-111">string</span></span>     |   <span data-ttu-id="09674-112">Microsoft Graph に対するアクセス許可の名前。たとえば、Files.Read.All です。</span><span class="sxs-lookup"><span data-stu-id="09674-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="09674-113">例</span><span class="sxs-lookup"><span data-stu-id="09674-113">Example</span></span>

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
