const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Dossier où sont stockés les fichiers JSON
const jsonFolder = './groupes'; // adapte le chemin

// Fonction pour lire JSON
function readJSON(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
}

function formatLdapTimestamp(ts) {
    if (!ts || typeof ts !== 'string') return '';
  
    // Nettoyer `.0Z` à la fin
    const cleanTs = ts.replace('.0Z', '');
  
    const year = cleanTs.substring(0, 4);
    const month = cleanTs.substring(4, 6);
    const day = cleanTs.substring(6, 8);
    const hour = cleanTs.substring(8, 10);
    const minute = cleanTs.substring(10, 12);
    const second = cleanTs.substring(12, 14);
  
    // Construire un objet Date
    const date = new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}Z`);
  
    // Formater en français
    return date.toLocaleString('fr-FR');
  }

// Fusionne les utilisateurs et ajoute la liste des groupes
function mergeUsersWithGroups(files) {
  const usersMap = new Map(); // key = user id, value = user object + groupes array

  files.forEach(file => {
    const groupName = path.basename(file, '.json');
    const users = readJSON(path.join(jsonFolder, file));

    users.forEach(user => {
      const userId = user.id;

      if (!usersMap.has(userId)) {
        // Clone l'objet utilisateur et ajoute un tableau groupes
        usersMap.set(userId, { ...user, groupes: [groupName] });
      } else {
        // Ajoute le groupe à l'utilisateur existant (sans doublon)
        const existingUser = usersMap.get(userId);
        if (!existingUser.groupes.includes(groupName)) {
          existingUser.groupes.push(groupName);
        }
      }
    });
  });

  return Array.from(usersMap.values());
}

// Convertit les utilisateurs pour Excel
function prepareForExcel(users) {
  return users.map(u => {
    const toDateString = (timestamp) =>
        timestamp ? new Date(timestamp).toLocaleString('fr-FR') : '';
    
    return {
        "Date de création": toDateString(u.createdTimestamp),
        "Nom d'utilisateur": u.username,
        "Email vérifié": u.emailVerified,
        "Nom": u.lastName,
        "Prénom": u.firstName,
        "Adresse email": u.email,
        "Statut": u.enabled ? 'actif' : 'désactivé',
        "Direction": u.attributes?.direction?.[0] || '',
        groupes: u.groupes.join(', ')

        // totp: u.totp,
        // federationLink: u.federationLink || '',  
        // Attributs LDAP individuels
        // LDAP_ENTRY_DN: u.attributes?.LDAP_ENTRY_DN?.[0] || '',
        // LDAP_ID: u.attributes?.LDAP_ID?.[0] || '',  
        // disableableCredentialTypes: (u.disableableCredentialTypes || []).join(', '),
        // requiredActions: (u.requiredActions || []).join(', '),
        // notBefore: u.notBefore,
        // "Date de création LDAP": formatLdapTimestamp(u.attributes?.createTimestamp?.[0] || ''),
        // "Date de modification LDAP": formatLdapTimestamp(u.attributes?.modifyTimestamp?.[0] || ''),
  
    };
  });
}

function main() {
  const files = fs.readdirSync(jsonFolder).filter(f => f.endsWith('.json'));

  const mergedUsers = mergeUsersWithGroups(files);
  const excelData = prepareForExcel(mergedUsers);

  // Crée une nouvelle feuille
  const worksheet = xlsx.utils.json_to_sheet(excelData);

  // Crée un nouveau classeur et ajoute la feuille
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Users');

  // Sauvegarde le fichier Excel
  const outputPath = 'users_realm_random.xlsx';
  xlsx.writeFile(workbook, outputPath);

  console.log(`✅ Fichier Excel généré : ${outputPath}`);
}

main();
