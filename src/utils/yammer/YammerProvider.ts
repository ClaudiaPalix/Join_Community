import axios from "axios";
import { IYammerProvider } from "./IYammerProvider";

export default class YammerProvider implements IYammerProvider {
  private readonly _apiUrl: string = "https://api.yammer.com/api/v1/";

  constructor(private aadToken: string, private currentUser: string) {}

  public async getGroups() {
    const userId = await this.getUserId(this.currentUser);
    const reqHeaders = {
      "content-type": "application/json",
      Authorization: `Bearer ${this.aadToken}`,
    };

    return axios.get(`${this._apiUrl}groups/for_user/${userId}`, {
      headers: reqHeaders,
    });
  }

  private async getUserId(email: string) {
    const reqHeaders = {
      Authorization: `Bearer ${this.aadToken}`,
    };
    const result = await axios.get(
      `${this._apiUrl}users/by_email.json?email=${email}`,
      { headers: reqHeaders }
    );
    return result.data[0].id;
  }
}
