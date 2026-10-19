// Muir Receipts — deployment configuration.
// Fill in the two Entra values after IT registers the app (see README.md).
// The SharePoint values below are the live Finance folder as at 19 Oct 2026.
window.MUIR_CONFIG = {
  // --- Microsoft Entra (Azure AD) app registration: single-page application ---
  clientId: "c20afc2d-76c9-4f0b-b2e7-a0b42ceafdb2",
  tenantId: "32f2aaea-cdbf-487a-ad81-3be35b773951",

  // --- Where receipts land: MuirEngineeringProjects › 9_Admin › Finance › Credit Card Receipts › Incoming ---
  driveId: "b!u3WZ9JZxBkCZDOWvPKaickgQY4Px_fhCsUnn--jhqzjYm4AYr5NrQJInxLG0rNKl",
  incomingFolderId: "01NLAK5FGOS3ETIYLXMZDZFQDY2VUCQF2Q",
  // Optional: a JSON file in the Credit Card Receipts folder that overrides the lists below
  // without redeploying the app (same shape as `defaults`). Leave null to use defaults only.
  configFileId: null,

  // Delegated Graph scopes the app asks for. Files.ReadWrite.All lets the user write to any
  // library they already have permission on — cardholders only get Contribute on `Incoming`.
  scopes: ["User.Read", "Files.ReadWrite.All"],

  // Image handling: long edge in px and JPEG quality. ~1600px keeps thermal-receipt text legible
  // and the file well under Graph's 4 MB simple-upload limit.
  image: { maxEdge: 1600, quality: 0.85 },

  defaults: {
    expenseTypes: [
      "Travel-Accommodation", "Travel-Flights", "Travel-Meals", "Travel-Ground",
      "Entertainment", "Motor Vehicle-Fuel", "Software/Subscriptions", "Telephone/Internet",
      "Tools & Consumables", "Freight/Postage", "Office", "Marketing/Trade Show", "Other"
    ],
    // Types that prompt for attendees (FBT): keep in step with the register's rules.
    attendeeTypes: ["Entertainment", "Travel-Meals"],
    currencies: ["AUD", "EUR", "USD", "GBP", "NZD", "SGD", "JPY", "CNY", "HKD", "AED", "Other"],
    projects: [],              // optional short list, e.g. ["MY Kalm", "HN834", "METS 2026"]
    reasonPlaceholder: "What it's for, and who was there if paying for others"
  }
};
