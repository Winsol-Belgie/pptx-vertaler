// Hulpscriptje (overgenomen van Feedbackloop) om lokaal een wachtwoord-hash
// te berekenen op EXACT dezelfde manier als de Worker (sha256(wachtwoord + pepper)).
// Nodig om de allereerste admin aan te maken; daarna beheer je gebruikers
// via de admin-sectie in de tool (Users-Excel).
//
// Gebruik:
//   node scripts/hash-password.js <wachtwoord> <pepper>
//
// <pepper> moet exact dezelfde waarde zijn als bij:
//   npx wrangler secret put PASSWORD_PEPPER

const crypto = require('crypto');

const [, , password, pepper] = process.argv;

if (!password || !pepper) {
  console.error('Gebruik: node scripts/hash-password.js <wachtwoord> <pepper>');
  process.exit(1);
}

console.log(crypto.createHash('sha256').update(password + pepper).digest('hex'));
