---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: b6cf82776f683929845df83c642b28ad024d665a
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596733"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="25df5-102">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="25df5-102">WebApplicationInfo element</span></span>

<span data-ttu-id="25df5-103">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="25df5-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="25df5-104">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="25df5-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="25df5-105">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="25df5-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="25df5-106">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="25df5-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="25df5-107">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](../requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25df5-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="25df5-108">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="25df5-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="25df5-109">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="25df5-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="25df5-110">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="25df5-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="25df5-111">子要素</span><span class="sxs-lookup"><span data-stu-id="25df5-111">Child elements</span></span>

|  <span data-ttu-id="25df5-112">要素</span><span class="sxs-lookup"><span data-stu-id="25df5-112">Element</span></span> |  <span data-ttu-id="25df5-113">必須</span><span class="sxs-lookup"><span data-stu-id="25df5-113">Required</span></span>  |  <span data-ttu-id="25df5-114">説明</span><span class="sxs-lookup"><span data-stu-id="25df5-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="25df5-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="25df5-115">**Id**</span></span>    |  <span data-ttu-id="25df5-116">必要</span><span class="sxs-lookup"><span data-stu-id="25df5-116">Yes</span></span>   |  <span data-ttu-id="25df5-117">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="25df5-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="25df5-118">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="25df5-118">**MsaId**</span></span>    |  <span data-ttu-id="25df5-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="25df5-119">No</span></span>   |  <span data-ttu-id="25df5-120">Msm.live.com に登録されている、アドインの web アプリケーションのクライアント ID。</span><span class="sxs-lookup"><span data-stu-id="25df5-120">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="25df5-121">**Resource**</span><span class="sxs-lookup"><span data-stu-id="25df5-121">**Resource**</span></span>  |  <span data-ttu-id="25df5-122">はい</span><span class="sxs-lookup"><span data-stu-id="25df5-122">Yes</span></span>   |  <span data-ttu-id="25df5-123">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="25df5-123">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="25df5-124">Scope</span><span class="sxs-lookup"><span data-stu-id="25df5-124">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="25df5-125">必要</span><span class="sxs-lookup"><span data-stu-id="25df5-125">Yes</span></span>  |  <span data-ttu-id="25df5-126">Microsoft Graph などのリソースに対してアドインが必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="25df5-126">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="25df5-127">Authorizations</span><span class="sxs-lookup"><span data-stu-id="25df5-127">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="25df5-128">いいえ</span><span class="sxs-lookup"><span data-stu-id="25df5-128">No</span></span>   | <span data-ttu-id="25df5-129">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="25df5-129">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="25df5-130">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="25df5-130">WebApplicationInfo example</span></span>

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
