// Sample decks for the dev-only /dev-deck command (moved from the POC 1 spike).
// typical  = realistic B1 English deck
// max      = German, long compounds, at the backend schema limits
// overflow = deliberately over the limits, to exercise shrink → reflow → split

export const FIXTURES = {
    typical: {
        title: 'Present Simple: Negatives',
        slides: [
            {
                layout: 'bullets',
                'slide-title': 'When do we use it?',
                items: [
                    'Facts that are always true',
                    'Habits and daily routines',
                    "Things we don't like or don't do",
                ],
            },
            {
                layout: 'pattern',
                'slide-title': 'How to form it',
                parts: ['Subject', "don't / doesn't", 'base verb'],
                example: { text: "She doesn't drink coffee.", highlight: "doesn't" },
            },
            {
                layout: 'comparison',
                'slide-title': 'Correct or wrong?',
                left: {
                    heading: 'Correct',
                    polarity: 'positive',
                    items: [
                        { text: "She doesn't like tea.", highlight: "doesn't" },
                        { text: "They don't play tennis.", highlight: "don't" },
                    ],
                },
                right: {
                    heading: 'Wrong',
                    polarity: 'negative',
                    items: [
                        { text: "She don't like tea.", highlight: "don't" },
                        { text: "They doesn't play tennis.", highlight: "doesn't" },
                    ],
                },
            },
            {
                layout: 'vocab-table',
                'slide-title': 'Food vocabulary',
                columns: ['Word', 'Meaning', 'Example'],
                rows: [
                    ['bread', 'food made from flour', 'I buy bread every day.'],
                    ['cheese', 'food made from milk', "He doesn't eat cheese."],
                    ['rice', 'small white grains', 'We cook rice for lunch.'],
                    ['apple', 'a round fruit', 'She has an apple.'],
                    ['soup', 'hot liquid food', "I don't like cold soup."],
                ],
            },
            {
                layout: 'examples',
                'slide-title': 'Examples',
                items: [
                    { text: "I don't work on Sundays.", highlight: "don't", translation: null },
                    { text: "My brother doesn't eat meat.", highlight: "doesn't", translation: null },
                    { text: "We don't watch TV in the morning.", highlight: "don't", translation: null },
                    { text: "It doesn't rain much in summer.", highlight: "doesn't", translation: null },
                ],
            },
        ],
    },

    max: {
        title: 'Trennbare Verben im Präsens: Bildung, Stellung und Ausnahmen',
        slides: [
            {
                layout: 'bullets',
                'slide-title': 'Wann benutzen wir trennbare Verben im Alltag?',
                items: [
                    'Das Präfix steht im Hauptsatz immer ganz am Satzende.',
                    'Im Nebensatz bleibt das Verb zusammen und steht am Ende.',
                    'Betonte Vorsilben wie an-, auf-, ein- sind meistens trennbar.',
                    'Unbetonte Vorsilben wie be-, ver-, zer- sind nie trennbar.',
                    'Im Perfekt steht -ge- zwischen Präfix und Verbstamm: eingekauft.',
                ],
            },
            {
                layout: 'pattern',
                'slide-title': 'Satzbau mit trennbaren Verben',
                parts: ['Subjekt', 'konjugiertes Verb', 'Ergänzungen', 'Präfix am Satzende'],
                example: { text: 'Die Lebensmittelverkäuferin räumt die Regale jeden Morgen ein.', highlight: 'räumt' },
            },
            {
                layout: 'comparison',
                'slide-title': 'Richtig oder falsch? Stellung des Präfixes',
                left: {
                    heading: 'Richtig',
                    polarity: 'positive',
                    items: [
                        { text: 'Ich kaufe am Samstag im Supermarkt ein.', highlight: 'ein' },
                        { text: 'Wir rufen unsere Großeltern heute Abend an.', highlight: 'an' },
                        { text: 'Der Zug fährt pünktlich um halb acht ab.', highlight: 'ab' },
                        { text: 'Sie räumt nach dem Frühstück die Küche auf.', highlight: 'auf' },
                    ],
                },
                right: {
                    heading: 'Falsch',
                    polarity: 'negative',
                    items: [
                        { text: 'Ich einkaufe am Samstag im Supermarkt.', highlight: 'einkaufe' },
                        { text: 'Wir anrufen unsere Großeltern heute Abend.', highlight: 'anrufen' },
                        { text: 'Der Zug abfährt pünktlich um halb acht.', highlight: 'abfährt' },
                        { text: 'Sie aufräumt nach dem Frühstück die Küche.', highlight: 'aufräumt' },
                    ],
                },
            },
            {
                layout: 'vocab-table',
                'slide-title': 'Wortschatz: Lebensmitteleinkauf',
                columns: ['Wort', 'Bedeutung', 'Beispielsatz'],
                rows: [
                    ['das Lebensmittelgeschäft', 'Laden für Essen und Getränke', 'Das Lebensmittelgeschäft öffnet um acht.'],
                    ['die Sonderangebotsabteilung', 'Bereich mit billigen Waren', 'Ich schaue zuerst in die Sonderangebotsabteilung.'],
                    ['der Einkaufswagen', 'Wagen für die Einkäufe', 'Der Einkaufswagen ist schon ganz voll.'],
                    ['die Mehrwegflasche', 'Flasche, die man zurückgibt', 'Bring bitte die Mehrwegflaschen zurück.'],
                    ['das Vollkornbrötchen', 'kleines Brot aus Vollkorn', 'Zum Frühstück esse ich ein Vollkornbrötchen.'],
                    ['die Kühltheke', 'gekühlter Verkaufstisch', 'Der Käse liegt in der Kühltheke.'],
                    ['der Kassenzettel', 'Beleg über den Einkauf', 'Hast du den Kassenzettel noch?'],
                    ['die Pfandrückgabe', 'Rückgabe von Pfandflaschen', 'Die Pfandrückgabe ist neben dem Eingang.'],
                ],
            },
            {
                layout: 'examples',
                'slide-title': 'Beispielsätze aus dem Alltag',
                items: [
                    { text: 'Ich kaufe jeden Freitag im Bioladen ein.', highlight: 'kaufe', translation: 'I shop at the organic store every Friday.' },
                    { text: 'Der Verkäufer packt die Einkäufe sorgfältig ein.', highlight: 'packt', translation: 'The salesman packs the groceries carefully.' },
                    { text: 'Wir probieren an der Käsetheke einen Bergkäse aus.', highlight: 'probieren', translation: 'We try a mountain cheese at the cheese counter.' },
                    { text: 'Meine Mitbewohnerin räumt den Kühlschrank auf.', highlight: 'räumt', translation: 'My flatmate tidies up the fridge.' },
                    { text: 'Die Kassiererin gibt mir das Wechselgeld zurück.', highlight: 'gibt', translation: 'The cashier gives me my change back.' },
                ],
            },
        ],
    },

    overflow: {
        title: 'Trennbare und untrennbare Verben im Präsens, Perfekt und Präteritum: der vollständige Überblick',
        slides: [
            {
                layout: 'bullets',
                'slide-title': 'Alle Regeln zu trennbaren und untrennbaren Verben auf einen Blick',
                items: [
                    'Das Präfix steht im Hauptsatz immer ganz am Satzende, auch wenn der Satz sehr lang ist.',
                    'Im Nebensatz bleibt das Verb zusammen und steht gemeinsam mit dem Präfix am Ende.',
                    'Betonte Vorsilben wie an-, auf-, aus-, ein-, mit-, vor- und zu- sind meistens trennbar.',
                    'Unbetonte Vorsilben wie be-, emp-, ent-, er-, ge-, miss-, ver- und zer- sind nie trennbar.',
                    'Im Perfekt steht -ge- zwischen Präfix und Verbstamm: eingekauft, angerufen, aufgeräumt.',
                    'Bei untrennbaren Verben fällt das -ge- im Partizip weg: bezahlt, verstanden, erzählt.',
                    'Einige Vorsilben wie über-, um-, unter- und durch- können trennbar oder untrennbar sein.',
                ],
            },
            {
                layout: 'pattern',
                'slide-title': 'Satzbau mit trennbaren Verben im Haupt- und Nebensatz',
                parts: ['Subjekt', 'konjugiertes Verb', 'Zeitangabe', 'Ortsangabe', 'Präfix am Satzende'],
                example: { text: 'Die Lebensmittelverkäuferin räumt die Regale jeden Morgen um sechs Uhr im Getränkemarkt ein.', highlight: 'räumt' },
            },
            {
                layout: 'comparison',
                'slide-title': 'Richtig oder falsch? Stellung des Präfixes im Satz',
                left: {
                    heading: 'Richtig',
                    polarity: 'positive',
                    items: [
                        { text: 'Ich kaufe am Samstagvormittag im großen Supermarkt ein.', highlight: 'ein' },
                        { text: 'Wir rufen unsere Großeltern heute Abend nach dem Essen an.', highlight: 'an' },
                        { text: 'Der Regionalzug fährt pünktlich um halb acht vom Hauptbahnhof ab.', highlight: 'ab' },
                        { text: 'Sie räumt nach dem Frühstück sofort die ganze Küche auf.', highlight: 'auf' },
                        { text: 'Die Kinder sehen am Wochenende gerne Zeichentrickfilme an.', highlight: 'an' },
                        { text: 'Mein Kollege bringt morgen selbstgebackenen Kuchen mit.', highlight: 'mit' },
                    ],
                },
                right: {
                    heading: 'Falsch',
                    polarity: 'negative',
                    items: [
                        { text: 'Ich einkaufe am Samstagvormittag im großen Supermarkt.', highlight: 'einkaufe' },
                        { text: 'Wir anrufen unsere Großeltern heute Abend nach dem Essen.', highlight: 'anrufen' },
                        { text: 'Der Regionalzug abfährt pünktlich um halb acht vom Hauptbahnhof.', highlight: 'abfährt' },
                        { text: 'Sie aufräumt nach dem Frühstück sofort die ganze Küche.', highlight: 'aufräumt' },
                        { text: 'Die Kinder ansehen am Wochenende gerne Zeichentrickfilme.', highlight: 'ansehen' },
                        { text: 'Mein Kollege mitbringt morgen selbstgebackenen Kuchen.', highlight: 'mitbringt' },
                    ],
                },
            },
            {
                layout: 'vocab-table',
                'slide-title': 'Wortschatz: Lebensmitteleinkauf (vollständige Liste)',
                columns: ['Wort', 'Bedeutung', 'Beispielsatz'],
                rows: [
                    ['das Lebensmittelgeschäft', 'Laden für Essen und Getränke', 'Das Lebensmittelgeschäft öffnet um acht.'],
                    ['die Sonderangebotsabteilung', 'Bereich mit billigen Waren', 'Ich schaue zuerst in die Sonderangebotsabteilung.'],
                    ['der Einkaufswagen', 'Wagen für die Einkäufe', 'Der Einkaufswagen ist schon ganz voll.'],
                    ['die Mehrwegflasche', 'Flasche, die man zurückgibt', 'Bring bitte die Mehrwegflaschen zurück.'],
                    ['das Vollkornbrötchen', 'kleines Brot aus Vollkorn', 'Zum Frühstück esse ich ein Vollkornbrötchen.'],
                    ['die Kühltheke', 'gekühlter Verkaufstisch', 'Der Käse liegt in der Kühltheke.'],
                    ['der Kassenzettel', 'Beleg über den Einkauf', 'Hast du den Kassenzettel noch?'],
                    ['die Pfandrückgabe', 'Rückgabe von Pfandflaschen', 'Die Pfandrückgabe ist neben dem Eingang.'],
                    ['das Tiefkühlgemüse', 'gefrorenes Gemüse', 'Tiefkühlgemüse ist praktisch und schnell.'],
                    ['die Obstabteilung', 'Bereich für Obst', 'Die Obstabteilung ist gleich links.'],
                    ['der Wochenmarkt', 'Markt einmal pro Woche', 'Samstags gehen wir auf den Wochenmarkt.'],
                    ['die Haltbarkeitsangabe', 'Datum, bis wann es gut ist', 'Prüf bitte die Haltbarkeitsangabe auf der Packung.'],
                ],
            },
            {
                layout: 'examples',
                'slide-title': 'Beispielsätze aus dem Alltag',
                items: [
                    { text: 'Ich kaufe jeden Freitag nach der Arbeit im kleinen Bioladen um die Ecke ein.', highlight: 'kaufe', translation: 'Every Friday after work I shop at the small organic store around the corner.' },
                    { text: 'Der Verkäufer packt die Einkäufe sorgfältig in die mitgebrachte Stofftasche ein.', highlight: 'packt', translation: 'The salesman carefully packs the groceries into the cloth bag we brought.' },
                    { text: 'Wir probieren an der Käsetheke einen würzigen Bergkäse aus.', highlight: 'probieren', translation: 'We try a spicy mountain cheese at the cheese counter.' },
                    { text: 'Meine Mitbewohnerin räumt jeden Sonntag den Kühlschrank auf.', highlight: 'räumt', translation: 'My flatmate tidies up the fridge every Sunday.' },
                    { text: 'Die Kassiererin gibt mir das Wechselgeld zurück.', highlight: 'gibt', translation: 'The cashier gives me my change back.' },
                    { text: 'Nach dem Einkauf bringen wir die leeren Pfandflaschen zurück.', highlight: 'bringen', translation: 'After shopping we bring the empty deposit bottles back.' },
                    { text: 'Am Wochenende laden wir unsere Nachbarn zum Grillen ein.', highlight: 'laden', translation: 'At the weekend we invite our neighbours for a barbecue.' },
                    { text: 'Der Bäcker stellt die frischen Brötchen um sechs Uhr bereit.', highlight: 'stellt', translation: 'The baker has the fresh rolls ready at six o’clock.' },
                ],
            },
        ],
    },
};

