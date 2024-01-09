const axios = require("axios");

const tenantId = "";
const clientId = "";
const clientSecret = "";
const resourceUrl = "https://graph.microsoft.com/";
const siteId = "";

const getAccessToken = async () => {
  let data = new FormData();
  data.append("grant_type", "client_credentials");
  data.append("client_id", clientId);
  data.append("client_secret", clientSecret);
  data.append("resource", resourceUrl);

  let config = {
    method: "post",
    maxBodyLength: Infinity,
    url: `https://login.microsoftonline.com/${tenantId}/oauth2/token`,
    headers: {
      ...data.getHeaders(),
    },
    data: data,
  };
  const response = await axios(config);
  return response.data.access_token;
};

const getSharePointData = async (accessToken) => {
  let config = {
    method: "get",
    maxBodyLength: Infinity,
    url: `https://graph.microsoft.com/v1.0/sites/${siteId}/analytics/lastSevenDays`,
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  };
  const response = await axios(config);
  console.log(response.data);
};

const main = async () => {
  const accessToken = await getAccessToken();
  //console.log(accessToken);
  await getSharePointData(accessToken);
};

main();
