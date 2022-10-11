import { UserInfo, UserManager } from '@hcllabs/openclientkeepcomponent';
import { Request } from '@loopback/rest';

export class UserContext implements UserInfo {
    public userId: string;
    public password?: string;
    public subject?: string;

  /**
   * Returns the userid and password passed on the request's authorization header when basic auth is being used. 
   * @param request The request being processed.
   * @returns the userid and password or undefined if basic auth is not used or the request does not contain an authorization header. 
   */
    static getUserInfo(request: Request): UserInfo {
        const userInfo = new UserContext();

        const authInfo = request.header("authorization")?.split(" ");
        if (authInfo !== undefined) {
            if (authInfo.length > 1) {
                if (authInfo[0] === "Basic") {
                    const buffer = Buffer.from(authInfo[1], "base64");
                    const unencoded = buffer.toString("utf-8");
                    const cred = unencoded.split(":");
                    userInfo.userId = cred[0];
                    userInfo.password = cred[1];
                }
            }
        }

        return userInfo;
    }

  /**
   * Returns the Keep X509 subject for the current user. 
   * @param request The request being processed.
   * @returns The X509 subject or undefined if the current user has not authenticated with Keep. 
   */
  static getSubjectFromRequest(request: Request): string | undefined {
    const userInfo = UserContext.getUserInfo(request);
    if (userInfo) {
      return UserManager.getInstance().getUserSubject(userInfo);
    }
    return undefined;
  }
}