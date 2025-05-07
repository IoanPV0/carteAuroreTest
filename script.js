const map = L.map('map').setView([46.5, 2.5], 6); // centre France
const apiKey = '2164e0e6e2b54796be75d32f6cba57f1';


L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap contributors'
}).addTo(map);


let cache = {};

fetch('geo_cache.json')
    .then(response => response.json())
    .then(data => {
    cache = data;
    })
    .catch(err => {
    console.warn("Pas de cache trouvé, un nouveau sera créé.");
});


window.onload = async function() { 
    try {
    const response = await fetch('Export_Office_20250428_141644_Export_cartographie_missions_et_signature_mail_165241.xls');
    const data = await response.arrayBuffer();

    const workbook = XLSX.read(data, { type: 'array' });

    map.eachLayer(layer => {
        if (!layer._url) map.removeLayer(layer);
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    const ouvertData = jsonData.filter(row => row.Statut === "Ouvert");
    console.clear();

    for (const row of ouvertData) {
        const adresse = `${row.Adresse}, ${row.CP} ${row.Ville}`;
        const territoire = `${row.Territoire}`;
        const nom = row["Nom de la mission"] || "Inconnu";
        const description = row["Description et objectifs"] || "Aucune description fournie";
        let dateOuverture = row["Date ouverture"];

        if (typeof dateOuverture === 'number') {
        dateOuverture = excelDateToJSDate(dateOuverture);
        }
        try {
        const coords = await getOrGeocode(adresse);
        if (coords) {
            L.marker([coords.lat, coords.lon]).addTo(map)
            .bindPopup(`
                <strong>${nom}</strong>
                <p>${adresse}</p>
                <p>Territoire : ${territoire}</p>
                <p>Date ouverture : ${dateOuverture}</p>
                <p>Description : ${description}</p>
                <button onclick="copyToClipboard(\`${sanitize(territoire)}\`, \`${sanitize(row.Adresse)}\`, \`${sanitize(row.Ville)}\`, \`${sanitize(row.CP)}\`)">Copier les informations dans le presse papier</button>
            `);
       }
        } catch (err) {
        console.error("Erreur de géocodage:", err);
        }
    }

    // Mettre à jour le cache dans le navigateur
    localStorage.setItem('geoCache', JSON.stringify(cache));
    exportCache();

    } catch (error) {
    console.error('Erreur lors de la lecture du fichier Excel:', error);
    }
};

async function getOrGeocode(adresse) {
    if (cache[adresse]) return cache[adresse];

    const url = `https://api.opencagedata.com/geocode/v1/json?q=${encodeURIComponent(adresse)}&key=${apiKey}&limit=1&no_annotations=1`;
    const res = await fetch(url);
    const data = await res.json();

    if (data.results.length > 0) {
    const coords = {
        lat: data.results[0].geometry.lat,
        lon: data.results[0].geometry.lng
    };
    cache[adresse] = coords;
    return coords;
    }

    return null;
}

function excelDateToJSDate(serial) {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info.toLocaleDateString('fr-FR');
}

function exportCache() {
    fetch('save_cache.php', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(cache)
    })
    .then(response => response.text())
    .then(data => {
    alert('Cache sauvegardé sur le serveur !');
    console.log(data);
    })
    .catch(error => {
    console.error('Erreur lors de la sauvegarde du cache :', error);
    });
}

function sanitize(str) {
    return String(str || '').replace(/`/g, "'").replace(/\\/g, "\\\\").replace(/\$/g, "\\$");
}


function copyToClipboard(territoire, adresse, ville, cp) {
    if (territoire === undefined)
    {
        territoire = '';
    }
    if (adresse === undefined)
    {
        adresse = '';
    }
    if (ville === undefined)
    {
        ville = '';
    }
    if (cp === undefined)
    {
        cp = '';
    }

    if (territoire.includes('SOLIFO'))
    {
        territoire = 'SOLIFO';
    }
    if (territoire.includes('Ouest'))
    {
        territoire = 'Ouest';
    }
    if (territoire.includes('SAVAH'))
    {
        territoire = 'SAVAH';
    }
    if (territoire.includes('Siège'))
    {
        territoire = 'Siège';
    }
    if (territoire.includes('PEIFS'))
    {
        territoire = 'PEIFS';
    }
    if (territoire.includes('PSC'))
    {
        territoire = 'PSC';
    }
    if (territoire.includes('SARAO'))
    {
        territoire = 'SARAO';
    }
    if (territoire.includes('PN'))
    {
        territoire = 'PN';
    }

    const adresseParts = [
        adresse || '',
        ville || '',
        cp || ''
    ]
    const adresseComplete = adresseParts.filter(Boolean).join(' ');

    const parts = [
        territoire || '',
        adresseComplete
    ];

    const text = parts.filter(Boolean).join('|');

    navigator.clipboard.writeText(text)
        .then(() => {
            alert('Informations copiées dans le presse-papiers !');
        })
        .catch(err => {
            console.error('Erreur lors de la copie :', err);
            alert('Erreur lors de la copie.');
        });
}
