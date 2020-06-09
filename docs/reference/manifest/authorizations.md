---
title: マニフェストファイルの承認要素
description: アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 675585f99fc6261a2145219d553f02b9f9abded3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608755"
---
# <a name="authorizations-element"></a><span data-ttu-id="db8a0-103">承認要素</span><span class="sxs-lookup"><span data-stu-id="db8a0-103">Authorizations element</span></span>

<span data-ttu-id="db8a0-104">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="db8a0-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="db8a0-105">**承認**は、マニフェスト内の[webapplicationinfo](webapplicationinfo.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="db8a0-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="db8a0-106">子要素</span><span class="sxs-lookup"><span data-stu-id="db8a0-106">Child elements</span></span>

|  <span data-ttu-id="db8a0-107">要素</span><span class="sxs-lookup"><span data-stu-id="db8a0-107">Element</span></span> |  <span data-ttu-id="db8a0-108">必須</span><span class="sxs-lookup"><span data-stu-id="db8a0-108">Required</span></span>  |  <span data-ttu-id="db8a0-109">説明</span><span class="sxs-lookup"><span data-stu-id="db8a0-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="db8a0-110">Authorization</span><span class="sxs-lookup"><span data-stu-id="db8a0-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="db8a0-111">はい</span><span class="sxs-lookup"><span data-stu-id="db8a0-111">Yes</span></span>     |   <span data-ttu-id="db8a0-112">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なスコープ (アクセス許可) を識別します。</span><span class="sxs-lookup"><span data-stu-id="db8a0-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="db8a0-113">例</span><span class="sxs-lookup"><span data-stu-id="db8a0-113">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
