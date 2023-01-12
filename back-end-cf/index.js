const EXPOSE_PATH = "";
const ONEDRIVE_REFRESHTOKEN = "";
const PASSWD_FILENAME = ".password";
const clientId = "78d4dc35-7e46-42c6-9023-2d39314433a5";
const clientSecret = "ZudGl-p.m=LMmr3VrKgAyOf-WevB3p50";
const loginHost = "https://login.microsoftonline.com";
const apiHost = "https://graph.microsoft.com";
const redirectUri = "http://localhost/onedrive-login";

addEventListener('scheduled', event => {
  event.waitUntil(fetchAccessToken(event.scheduledTime));
});

addEventListener("fetch", (event) => {
  try {
    return event.respondWith(handleRequest(event.request));
  } catch (e) {
    return event.respondWith(new Response("Error thrown " + e.message));
  }
});

const OAUTH = {
  redirectUri: redirectUri,
  refreshToken: ONEDRIVE_REFRESHTOKEN,
  clientId: clientId,
  clientSecret: clientSecret,
  oauthUrl: loginHost + "/common/oauth2/v2.0/",
  apiUrl: apiHost + "/v1.0/me/drive/root",
  scope: apiHost + "/Files.ReadWrite.All offline_access",
};

async function handleRequest(request) {
  let querySplited, requestPath;
  let queryString = decodeURIComponent(request.url.split("?")[1]);
  if (queryString) querySplited = queryString.split("=");
  if (querySplited && querySplited[0] === "file") {
    const file = querySplited[1];
    const fileName = file.split("/").pop();
    if (fileName === PASSWD_FILENAME)
      return Response.redirect(
        "https://www.baidu.com/s?wd=%E6%80%8E%E6%A0%B7%E7%9B%97%E5%8F%96%E5%AF%86%E7%A0%81",
        301
      );
    requestPath = file.replace("/" + fileName, "");
    const url = await fetchFiles(requestPath, fileName);
    return Response.redirect(url, 302);
  } else {
    const { headers } = request;
    const contentType = headers.get("content-type");
    let body = {};
    if (contentType && contentType.includes("form")) {
      const formData = await request.formData();
      for (let entry of formData.entries()) {
        body[entry[0]] = entry[1];
      }
    }
    requestPath = Object.getOwnPropertyNames(body).length ? body["?path"] : "";
    const files = await fetchFiles(requestPath, null, body.passwd);
    const getFolderListResponse = new Response(files, {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Cache-Control": "max-age=3600",
        "Content-Type": "application/json; charset=utf-8",
      },
    });
    console.log(getFolderListResponse)
    return getFolderListResponse;
  }
}

async function gatherResponse(response) {
  const { headers } = response;
  const contentType = headers.get("content-type");
  if (contentType.includes("application/json")) {
    return await response.json();
  } else if (contentType.includes("application/text")) {
    return await response.text();
  } else if (contentType.includes("text/html")) {
    return await response.text();
  } else {
    return await response.text();
  }
}

async function cacheFetch(url, options) {
  return fetch(new Request(url, options), {
    cf: {
      cacheTtl: 3600,
      cacheEverything: true,
    },
  });
}

async function getContent(url) {
  const response = await cacheFetch(url);
  const result = await gatherResponse(response);
  return result;
}

async function getContentWithHeaders(url, headers) {
  const response = await cacheFetch(url, { headers: headers });
  const result = await gatherResponse(response);
  return result;
}

async function fetchFormData(url, data) {
  const formdata = new FormData();
  for (const key in data) {
    if (data.hasOwnProperty(key)) {
      formdata.append(key, data[key]);
    }
  }
  const requestOptions = {
    method: "POST",
    body: formdata,
  };
  const response = await cacheFetch(url, requestOptions);
  const result = await gatherResponse(response);
  return result;
}

async function fetchAccessToken() {
  let refreshToken = OAUTH["refreshToken"];
  if (typeof FODI_CACHE !== 'undefined') {
    const cache = JSON.parse(await FODI_CACHE.get('token_data'));
    if (cache?.refresh_token) {
      const passedMilis = Date.now() - cache.save_time;
      if (passedMilis / 1000 < cache.expires_in - 600) {
        return cache.access_token;
      }

      if (passedMilis < 6912000000) {
        refreshToken = cache.refresh_token;
      }
    }
  }

  const url = OAUTH["oauthUrl"] + "token";
  const data = {
    client_id: OAUTH["clientId"],
    client_secret: OAUTH["clientSecret"],
    grant_type: "refresh_token",
    requested_token_use: "on_behalf_of",
    refresh_token: refreshToken,
  };
  const result = await fetchFormData(url, data);

  if (typeof FODI_CACHE !== 'undefined' && result?.refresh_token) {
    result.save_time = Date.now();
    await FODI_CACHE.put('token_data', JSON.stringify(result));
  }
  return result.access_token;
}


async function filerCaller(path, accessToken) {

  const uri =
      OAUTH.apiUrl +
      encodeURI(path) +
      "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
  return await getContentWithHeaders(uri, {
    Authorization: "Bearer " + accessToken,
  });

}

async function filerCallerByDriveIdAndItemId(driveId, itemId, accessToken) {
  const rUrl = apiHost + "/v1.0/drives/" + rDId + "/items/" + rId + "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
  return await getContentWithHeaders(rUrl, {
    Authorization: "Bearer " + accessToken,
  });
  
}
async function filerCallerBySharedPath(drivePath, accessToken) {
  const rUrl = apiHost + "/v1.0/drives" + drivePath + "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
  return await getContentWithHeaders(rUrl, {
    Authorization: "Bearer " + accessToken,
  });
  
}

async function fetchFiles(path, fileName, passwd) {
  if (path === "/") path = "";
  let pathPre = path
  let paths = path.split('/').filter(n => n)
  if (paths.length === 0) {
    paths = [""]
  }
  let remoteFolder = false
  let body = null
  const accessToken = await fetchAccessToken();

  let isSharedFolder = false
  let concatPath = ""
  let sharedPath = ""
  let normalPath = ""
  for (let levlelPath in paths) {
    concatPath = "/" + paths[levlelPath]
    
    if ((!normalPath && !sharedPath) && (paths[levlelPath] || EXPOSE_PATH)) paths[levlelPath] = ":/" + EXPOSE_PATH + paths[levlelPath];
    let uri
    if (!isSharedFolder) {
      let tempPath = paths[levlelPath]
      if (normalPath) {
        tempPath = ":" + normalPath + "/" + paths[levlelPath]
      }
      uri =
      OAUTH.apiUrl +
      encodeURI(tempPath) +
      "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
      body = await getContentWithHeaders(uri, {
        Authorization: "Bearer " + accessToken,
      });

    } else {
      let tempPath =  sharedPath  + "/" + paths[levlelPath] 
      uri = apiHost + "/v1.0" + encodeURI(tempPath) + "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
      body = await getContentWithHeaders(uri, {
        Authorization: "Bearer " + accessToken,
      });
    }
    
    if (body && body.remoteItem) {
      isSharedFolder = true
      const rDId = body.remoteItem.parentReference.driveId
      const rId = body.remoteItem.id
      uri = apiHost + "/v1.0/drives/" + rDId + "/items/" + rId + "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
      body = await getContentWithHeaders(uri, {
        Authorization: "Bearer " + accessToken,
      });
      if (body && body.children && body.children.length > 0) {
        sharedPath = body.children[0].parentReference.path
        sharedPath = decodeURI(sharedPath)
        
      }
    }else{
      if (body && body.children && body.children[0] && body.children[0].parentReference&&body.children[0].parentReference.path  ){
          normalPath = normalPath + body.children[0].parentReference&&body.children[0].parentReference.path.split(":")[1]
          normalPath = decodeURI(normalPath)
      }
    }
  }
  if (fileName) {
    let thisFile = null;
    body.children.forEach((file) => {
      if (file.name === decodeURIComponent(fileName)) {
        thisFile = file["@microsoft.graph.downloadUrl"];
        return;
      }
    });
    return thisFile;
  } else {
    let files = [];
    let sharedfiles = [];
    let encrypted = false;
    for (let i = 0; i < body.children.length; i++) {
      const file = body.children[i];
      if (file.remoteItem) {
        const remoteItemDriveId = file.remoteItem.parentReference.driveId
        const remoteItemId = file.remoteItem.id
        const retrieveShareFileUri = apiHost + "/v1.0/drives/" + remoteItemDriveId + "/root" + "?expand=children(select=name,size,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl,remoteItem)";
        const sharedFolderResBody = await getContentWithHeaders(retrieveShareFileUri, {
          Authorization: "Bearer " + accessToken,
        });
        for (let i = 0; i < sharedFolderResBody.children.length; i++) {
          const sharedFile = sharedFolderResBody.children[i];
          sharedfiles.push({
            name: sharedFile.name,
            size: sharedFile.size,
            time: sharedFile.lastModifiedDateTime,
            url: sharedFile["@microsoft.graph.downloadUrl"],
          });
        }
        
      } else {
        files.push({
          name: file.name,
          size: file.size,
          time: file.lastModifiedDateTime,
          url: file["@microsoft.graph.downloadUrl"],
        });
      }
    }
    const allFiles = files.concat(files, sharedfiles)
    let parent = pathPre
    parent = pathPre==="" ? "/" : pathPre
    path = path.replace(":","")
    path= path==="" ? "/" : path
    if (encrypted) {
      return JSON.stringify({ parent: parent, files: [], encrypted: true , path: path , pathPre: pathPre});
    } else {
      return JSON.stringify({ parent: parent, files: [...files, ...sharedfiles] , path: path, pathPre: pathPre});
    }
  }
  
}
