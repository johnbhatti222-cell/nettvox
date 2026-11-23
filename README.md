# ISP Demarcation Map

Map of iON Network, Nettvox and Satellite Netcom subscribers per building, with a demarcation of who operates each building and zone, based only on subscriber volume.

- **Frontend:** `index.html`, hosted on GitHub Pages
- **Backend:** `code.gs`, in Google Apps Script, bound to the Google Sheet **ISP Confluence Map - Data** (tab `Data`)
- **Map:** Leaflet with free CARTO / OpenStreetMap / Esri tiles. No Google Maps API key is needed.
- **Geocoding:** Apps Script's built-in `Maps.newGeocoder()`. It is free (daily quota) and needs no key.

## Demarcation rules

| Item | Rule |
|---|---|
| Building operator | The ISP with the most subscribers in the building. |
| Tie in a building | The tied ISP that leads the building's zone gets it. If that is also tied, the area leader, then the overall leader. |
| Zone | A grid square (250 m / 500 m / 1 km / 2 km, selectable). It is coloured by the ISP with the most subscribers across all its buildings. Thick lines mark borders between different operators. |
| "Zone majority" switch | Every building in a zone goes to the zone leader, so each zone has a single field team. |
| Colours | Orange = iON Network, Blue = Satellite Netcom, Yellow = Nettvox |

The **Demarcation** tab shows, for each ISP:
- buildings and zones operated
- subscribers handed over and taken in
- subscribers after demarcation

It also shows a From → To transfer table. The **Buildings** tab can:
- download the full result as Excel (Summary, Buildings, Transfers, Zones)
- write it to a `Demarcation` tab in the Google Sheet

## Setup

1. **Apps Script**
   1. Open the Google Sheet and go to Extensions → Apps Script.
   2. Paste `code.gs` and save.
   3. In `setPasscode()`, change `'change-me'` to your passcode.
   4. Select `setPasscode` in the toolbar, click **Run**, and approve the permissions.
   5. Go to Deploy → New deployment → type **Web app**. Set Execute as **Me** and Who has access **Anyone**. Copy the `/exec` URL.
   6. When you edit `code.gs` later, go to Deploy → Manage deployments → ✏️ → Version **New version**. This keeps the same URL.
2. **GitHub Pages**
   1. In `index.html`, set `API_URL` (near the top of the script) to the `/exec` URL.
   2. Go to repo Settings → Pages → Deploy from branch → `main` / root.
   3. Open `https://<user>.github.io/nettvox/` and enter the passcode.

## Uploading data later

Go to **Data** tab → **Download Excel template**, fill it in, and drop it back in. You have two modes:
- **Add / update** (default): matches buildings on name + area + city and updates their counts.
- **Replace all**: clears the sheet and loads only the new file.

Latitude/longitude are optional. Buildings without them are placed approximately in their area. You can fix them later in **Fix locations**, or with **Locate next batch**, which uses the free geocoder.
