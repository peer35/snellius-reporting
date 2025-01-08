from ldap3 import Server, Connection, ALL, NTLM, ALL_ATTRIBUTES
from config import AD_PW, AD_USER, AD_SERVER, REPORT_PATH
import json

USERDATA_FILE = "data/userdata.json"  # store all found userdata here.


def ad_lookup(datafile, lookup=True):
    with open(datafile, "r+") as fp:
        data = json.load(fp)
    with open(f"{USERDATA_FILE}", "r+") as fp:
        userdata = json.load(fp)
    if lookup:
        server = Server(AD_SERVER, use_ssl=True, get_info=ALL)
        conn = Connection(server, AD_USER, AD_PW, auto_bind=True)
    for account, reportdata in data.items():
        email = reportdata.get("email", "")
        # we could try to add a lookup by first/last name for the other addresses later
        print(email)
        if email.endswith(("vu.nl", "acta.nl")) and lookup:  # VU users
            conn.search(
                "dc=vu,dc=local",
                f"(&(objectclass=person)(|(proxyaddresses=SMTP:{email})(proxyaddresses=smtp:{email})))",
                attributes=[
                    "department",
                    "company",
                    "eduPersonAffiliation",
                    "title",
                    "displayName",
                ],
            )
            try:
                entry = conn.entries[0]
                data[account]["AD"] = {
                    "department": "|".join(entry.department),
                    "company": "|".join(entry.company),
                    "eduPersonAffiliation": "|".join(entry.eduPersonAffiliation),
                    "title": "|".join(entry.title),
                    "displayName": "|".join(entry.displayName),
                    "account": account,
                }
                userdata[email] = data[account]["AD"]
            except IndexError:  # not found in AD
                # lookup in userdata
                if (
                    email in userdata
                ):  # use info found previously in case account has been deleted
                    data[account]["AD"] = userdata[email]
                else:
                    data[account]["AD"] = {}
        else:
            if (
                email in userdata
            ):  # possible to store manually connected data in the userdata.json
                data[account]["AD"] = userdata[email]
            else:
                data[account]["AD"] = {}

    with open(f"{USERDATA_FILE}", "w") as fp:
        json.dump(userdata, fp)
    ad_datafile = datafile.replace(".json", "_AD.json")
    with open(ad_datafile, "w") as fp:
        json.dump(data, fp)
    return ad_datafile


if __name__ == "__main__":
    datafile = f"{REPORT_PATH}/2307090_24.20250106.json"
    ad_datafile = ad_lookup(datafile)
