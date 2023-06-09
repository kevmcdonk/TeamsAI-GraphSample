import { AppCredential, OnBehalfOfUserCredential, createMicrosoftGraphClient, createMicrosoftGraphClientWithCredential } from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";

export class GraphService {
    private graphClient: Client;
    private _token: string;


    constructor(token: string) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this._token = token;
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    async createConnection(connection, connectorTicket: string) {
        await this.graphClient.api("/external/connections")
            .version("beta")
            .header("GraphConnectors-Ticket", connectorTicket)
            .post(connection);
    }

    async createSchema(connectionId: string, schema) {
        await this.graphClient.api(`/external/connections/${connectionId}/schema`)
            .version("beta")
            .post(schema);
    }

    async getConnection(connectionId: string) {
        return this.graphClient.api(`/external/connections/${connectionId}`)
            .version("beta")
            .get();
    }

    async createExternalItem(connectionId: string, itemId: string, externalItem) {
        await this.graphClient.api(`/external/connections/${connectionId}/items/${itemId}`)
            .version("beta")
            .put(externalItem);
    }

    async deleteConnection(connectionId: string) {
        await this.graphClient.api(`/external/connections/${connectionId}`)
            .version("beta")
            .delete();
    }

    async getUsersMail() {
        const mails = await this.graphClient.api("/me/messages")
            .version("beta")
            .get();
        return mails;
    }

    async getSites(query: string) {
        const sites = await this.graphClient.api("/sites?search="+query)
            .version("beta")
            .get();
        return sites;
    }

    async sendMail(subject:string, recipient: string, body: string) {
        const sendMail = {
            message: {
              subject: subject,
              body: {
                contentType: 'Text',
                content: body
              },
              toRecipients: [
                {
                  emailAddress: {
                    address: recipient
                  }
                }
              ]
            },
            saveToSentItems: 'false'
          };
          
          await this.graphClient.api('/me/sendMail')
              .post(sendMail);
    }

    async searchFiles(searchQuery: string) {
        const searchPost = {
            requests: [
                {
                    entityTypes: [
                      "driveItem", "listItem", "list"
                    ],
                    query: {
                      queryString: searchQuery
                    },
                    "from": 0,
                    "size": 10
                  } 
            ]
        };

        const searchResults = await this.graphClient.api('/search/query')
              .post(searchPost);
        return searchResults;
    }
}

