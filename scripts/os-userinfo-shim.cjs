const os = require("node:os");
const original = os.userInfo;
os.userInfo = function patchedUserInfo(options) {
  try { return original(options); }
  catch (error) {
    if (error?.code !== "ERR_SYSTEM_ERROR") throw error;
    return { uid: -1, gid: -1, username: process.env.USERNAME || "site-builder",
      homedir: process.env.USERPROFILE || process.cwd(), shell: null };
  }
};
