/*
 * Playing cards for Solitaire, FreeCell and Hearts: the deck, the shuffle,
 * card faces drawn as SVG at the classic 71 x 96 pixels, card backs, and
 * drag-and-drop with the mouse or a finger.
 *
 * A card is { rank: 1..13, suit: "C" | "D" | "H" | "S" }.
 */
(function () {
  "use strict";

  var Win95 = (window.Win95 = window.Win95 || {});

  var W = 71;
  var H = 96;
  var RED = "#d40000";
  var SUITS = ["C", "D", "H", "S"];
  var RANKS = [null, "A", "2", "3", "4", "5", "6", "7", "8", "9", "10", "J", "Q", "K"];
  var RANK_NAMES = [null, "Ace", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Jack", "Queen", "King"];
  var SUIT_NAMES = { C: "Clubs", D: "Diamonds", H: "Hearts", S: "Spades" };

  // Each suit drawn in a 100 x 100 box.
  var SUIT_PATH = {
    H: "M50 92C30 72 4 54 4 30 4 13 16 4 28 4c10 0 18 6 22 16C54 10 62 4 72 4c12 0 24 9 24 26 0 24-26 42-46 62z",
    D: "M50 2L88 50 50 98 12 50z",
    S: "M50 2C62 24 94 40 94 64c0 14-10 23-22 23-8 0-14-4-18-10 1 9 5 16 12 21H34c7-5 11-12 12-21-4 6-10 10-18 10C16 87 6 78 6 64 6 40 38 24 50 2z",
    C:
      "M50 4a21 21 0 1 1 0 42 21 21 0 1 1 0-42z" +
      "M27 38a21 21 0 1 1 0 42 21 21 0 1 1 0-42z" +
      "M73 38a21 21 0 1 1 0 42 21 21 0 1 1 0-42z" +
      "M38 36h24l-6 24H44z" +
      "M46 56h8c0 18 4 30 14 40H32c10-10 14-22 14-40z",
  };

  function isRed(suit) {
    return suit === "H" || suit === "D";
  }

  function newDeck() {
    var deck = [];
    SUITS.forEach(function (suit) {
      for (var r = 1; r <= 13; r++) deck.push({ rank: r, suit: suit });
    });
    return deck;
  }

  // Fisher-Yates. `rand` returns a float in [0, 1); Math.random by default.
  function shuffle(arr, rand) {
    rand = rand || Math.random;
    for (var i = arr.length - 1; i > 0; i--) {
      var j = Math.floor(rand() * (i + 1));
      var t = arr[i];
      arr[i] = arr[j];
      arr[j] = t;
    }
    return arr;
  }

  function name(card) {
    return RANK_NAMES[card.rank] + " of " + SUIT_NAMES[card.suit];
  }

  function key(card) {
    return card.suit + card.rank;
  }

  // ---------------------------------------------------------------- faces

  function pip(suit, cx, cy, size, flip, fill) {
    var t =
      (flip ? "rotate(180 " + cx + " " + cy + ") " : "") +
      "translate(" + (cx - size / 2) + " " + (cy - size / 2) + ") scale(" + size / 100 + ")";
    return '<path d="' + SUIT_PATH[suit] + '" transform="' + t + '" fill="' + fill + '"/>';
  }

  var L = 21, C = 35.5, R = 50;
  var PIPS = {
    2: [[C, 17], [C, 79]],
    3: [[C, 17], [C, 48], [C, 79]],
    4: [[L, 17], [R, 17], [L, 79], [R, 79]],
    5: [[L, 17], [R, 17], [C, 48], [L, 79], [R, 79]],
    6: [[L, 17], [R, 17], [L, 48], [R, 48], [L, 79], [R, 79]],
    7: [[L, 17], [R, 17], [C, 32.5], [L, 48], [R, 48], [L, 79], [R, 79]],
    8: [[L, 17], [R, 17], [C, 32.5], [L, 48], [R, 48], [C, 63.5], [L, 79], [R, 79]],
    9: [[L, 17], [R, 17], [L, 37.7], [R, 37.7], [C, 48], [L, 58.3], [R, 58.3], [L, 79], [R, 79]],
    10: [[L, 17], [R, 17], [C, 27.3], [L, 37.7], [R, 37.7], [L, 58.3], [R, 58.3], [C, 68.7], [L, 79], [R, 79]],
  };

  // ---------------------------------------------------------------- courts
  //
  // Jack, Queen and King. Each is drawn once as the top half of the card --
  // head, shoulders, and whatever the figure is holding -- and the card shows
  // that half twice, the second copy turned through 180 degrees, the way a
  // double-ended court card is printed. The figures are drawn in a 46 x 38.5
  // window whose bottom edge is the middle of the card.
  //
  // At 71 x 96 a head is a dozen pixels wide, so the drawing is flat colour
  // in bold shapes with a thin outline; anything finer turns to grey.

  var INK = "#151515";
  var SKIN = "#f8d5ab";
  var GOLD = "#f6c40e";
  var BLUE = "#1d4fc6";
  var STEEL = "#dfe7f0";

  // What each house wears: the robe, the panel down its front, the trim, and
  // the colour of its hair.
  var COURT = {
    S: { robe: "#2b2b33", panel: RED, trim: GOLD, hair: "#4a3323" },
    H: { robe: RED, panel: BLUE, trim: GOLD, hair: "#c9781e" },
    D: { robe: GOLD, panel: RED, trim: BLUE, hair: "#f1cf5a" },
    C: { robe: BLUE, panel: GOLD, trim: RED, hair: "#6b4426" },
  };

  var OUT = 'stroke="' + INK + '" stroke-width="0.6" stroke-linejoin="round"';
  var FINE = 'fill="none" stroke="' + INK + '" stroke-width="0.5" stroke-linecap="round"';

  function shape(d, fill, extra) {
    return '<path d="' + d + '" fill="' + fill + '" ' + (extra || OUT) + "/>";
  }

  // Both sides of something drawn on the left.
  function bothSides(markup) {
    return markup + '<g transform="translate(71 0) scale(-1 1)">' + markup + "</g>";
  }

  function courtBody(rank, pal) {
    var out = [];
    // Shoulders and sleeves, with two bands of trim across each sleeve.
    out.push(shape("M12.5 48V41.5q1.5-7 11-8.8L30 31h11l6.5 1.7q9.5 1.8 11 8.8V48z", pal.robe));
    out.push(bothSides(shape("M12.5 43.5l9.8-10.2 2.3-.6L12.5 45.8zM12.5 48l14-14.6 1.8-.5L14.7 48z", pal.trim, 'stroke="none"')));
    // The panel down the front, and the lozenges on it.
    out.push(shape("M29.5 31h12l2.5 17h-17z", pal.panel));
    out.push(shape("M35.5 35.5l2 2.7-2 2.7-2-2.7zM35.5 41.3l2 2.7-2 2.7-2-2.7z", pal.trim, 'stroke="' + INK + '" stroke-width="0.35"'));
    out.push('<path d="M30.7 38.2l-1.4 9.8M40.3 38.2l1.4 9.8" ' + FINE + "/>");
    if (rank === 13) {
      // An ermine collar.
      out.push(shape("M25 33.5q10.5 9 21 0l-3.5-3q-7 5-14 0z", "#fff"));
      out.push('<path d="M28.5 34.2v1.3M32 36v1.3M35.5 36.6v1.3M39 36v1.3M42.5 34.2v1.3" stroke="' + INK + '" stroke-width="0.8" stroke-linecap="round"/>');
    } else if (rank === 12) {
      // An open neckline and a string of pearls.
      out.push(shape("M30 30.5q5.5 8 11 0z", SKIN));
      out.push('<g fill="#fff" stroke="' + INK + '" stroke-width="0.35"><circle cx="31.3" cy="33.3" r="0.9"/><circle cx="33.2" cy="35" r="0.9"/><circle cx="35.5" cy="35.7" r="0.9"/><circle cx="37.8" cy="35" r="0.9"/><circle cx="39.7" cy="33.3" r="0.9"/></g>');
    } else {
      // A turned-down collar.
      out.push(bothSides(shape("M35.5 34.5l-6.5-4-2.5 3.3 6.5 3.7z", "#fff")));
    }
    // The band where the two halves of the card meet.
    out.push(shape("M12.5 45.7h46V48h-46z", pal.trim, 'stroke="' + INK + '" stroke-width="0.5"'));
    return out.join("");
  }

  // The face itself, drawn about (0, 0) looking a little to the right; the
  // caller mirrors it for a figure that looks left.
  function courtFace(rank, pal) {
    var out = [];
    var beard = rank === 13;
    out.push('<path d="M-1.9-1.9h-2.2M1.3-1.9h2.3" ' + FINE + "/>");
    out.push('<ellipse cx="-2.6" cy="-0.5" rx="0.85" ry="0.75" fill="' + INK + '"/><ellipse cx="2.9" cy="-0.5" rx="0.85" ry="0.75" fill="' + INK + '"/>');
    out.push('<path d="M0.6-0.6l1 2.9-1.3.3" ' + FINE + "/>");
    if (rank === 12) {
      out.push('<circle cx="-3.4" cy="2" r="1.1" fill="#f4a08c" opacity="0.7"/><circle cx="3.9" cy="2" r="1" fill="#f4a08c" opacity="0.7"/>');
      out.push(shape("M-1.2 4.2q.9-.7 1.6-.1.7-.6 1.6.1-1.6 1.5-3.2 0z", RED, 'stroke="none"'));
    } else if (beard) {
      out.push(shape("M-3.8 3.7q2.3-1.5 4.2-.3 1.9-1.2 4.2.3-2.2 1.2-4.2.5-2 .7-4.2-.5z", "#fff", 'stroke="' + INK + '" stroke-width="0.4"'));
      out.push('<path d="M-.8 5.6h2.4" stroke="' + RED + '" stroke-width="0.8" stroke-linecap="round"/>');
    } else {
      out.push('<path d="M-1.2 4.1q1.5 1.1 3.1 0" fill="none" stroke="' + RED + '" stroke-width="0.8" stroke-linecap="round"/>');
    }
    return out.join("");
  }

  function courtHead(rank, pal, dir) {
    var out = [];
    var hair = pal.hair;
    if (rank === 13) {
      // Hair to the collar, then the head, then a beard over the jaw.
      out.push(bothSidesAt(shape("M-5-4.5q-4.5 2.5-4 7.5.3 3 2.7 3.5L-5 2z", hair)));
      out.push('<ellipse cx="0" cy="0" rx="6" ry="7" fill="' + SKIN + '" ' + OUT + "/>");
      out.push(shape("M-6.1.8q-1 7.7 3.3 10.9L0 9.6l2.8 2.1Q7.1 8.5 6.1.8 4.3 4.5 0 4.2-4.3 4.5-6.1.8z", hair));
    } else if (rank === 12) {
      // Long hair behind, and a fringe.
      out.push(bothSidesAt(shape("M-4-6q-6.5 2-5.8 9.5.4 4.5 2.3 7.5l4-1.5z", hair)));
      out.push(shape("M-2.2 5h4.4v5h-4.4z", SKIN, 'stroke="none"'));
      out.push('<ellipse cx="0" cy="0" rx="5.5" ry="6.7" fill="' + SKIN + '" ' + OUT + "/>");
      out.push(shape("M-5.6-1.5q.3-5.5 5.6-5.6 5.3.1 5.6 5.6-2.6-3.3-5.6-3.3t-5.6 3.3z", hair));
    } else {
      // A page's bob, curling out at the bottom.
      out.push(bothSidesAt(shape("M-4.5-5q-4.2 2.5-3.8 8 .2 3.5-1.2 4.5 3.5 1.2 5.3-1.5z", hair)));
      out.push('<ellipse cx="0" cy="0" rx="5.7" ry="6.9" fill="' + SKIN + '" ' + OUT + "/>");
      out.push(shape("M-5.8-1.8q2.5-3 5.8-3t5.8 3q-.3-4.5-5.8-4.7-5.5.2-5.8 4.7z", hair));
    }
    out.push('<g transform="scale(' + dir + ' 1)">' + courtFace(rank, pal) + "</g>");
    if (rank === 13) {
      // A tall crown: a jewelled band, five points, a cap inside it.
      out.push(shape("M-6-8.5h12l-.5-4h-11z", RED, 'stroke="none"'));
      out.push(shape("M-7-6v-2.8l-.8-5 3.5 3 1.9-3.9L0-10.8l2.4-3.9 1.9 3.9 3.5-3 -.8 5V-6z", GOLD));
      out.push('<path d="M-7-8.3H7" ' + FINE + "/>");
      out.push('<circle cx="-4" cy="-7.1" r="0.75" fill="' + RED + '"/><circle cx="0" cy="-7.1" r="0.75" fill="' + BLUE + '"/><circle cx="4" cy="-7.1" r="0.75" fill="' + RED + '"/>');
    } else if (rank === 12) {
      // A coronet with a pearl on each point.
      out.push(shape("M-5.5-5.8v-2.2l-.7-3.3 3.2 2.3L0-12.5 3-9l3.2-2.3-.7 3.3v2.2z", GOLD));
      out.push('<g fill="#fff" stroke="' + INK + '" stroke-width="0.35"><circle cx="-6.2" cy="-11.5" r="0.9"/><circle cx="0" cy="-12.9" r="0.9"/><circle cx="6.2" cy="-11.5" r="0.9"/></g>');
      out.push('<circle cx="0" cy="-7.3" r="0.8" fill="' + pal.panel + '" stroke="' + INK + '" stroke-width="0.3"/>');
    } else {
      // A soft cap with a band and a plume swept back from the way he looks.
      out.push(shape("M-7.3-4.5q-1.2-6.3 7.3-6.5 8.5.2 7.3 6.5-7.3-2.7-14.6 0z", pal.robe));
      out.push(shape("M-7.3-4.5q7.3-2.7 14.6 0l-.3-1.9q-7-2.5-14 0z", pal.trim, 'stroke="' + INK + '" stroke-width="0.4"'));
      out.push('<g transform="scale(' + -dir + ' 1)">' + shape("M3-9.5q3-5 7.5-3.3-1 .2-1.4 1.1 1.4-.2 1.9.8-4-.6-6 3.2z", "#fff") + "</g>");
    }
    return out.join("");
  }

  // The same, about the middle of the head rather than of the card.
  function bothSidesAt(markup) {
    return markup + '<g transform="scale(-1 1)">' + markup + "</g>";
  }

  // What the figure holds, and the hand that holds it.
  function courtProp(rank, suit, pal) {
    var hand = function (cx, cy) {
      return '<ellipse cx="' + cx + '" cy="' + cy + '" rx="2.6" ry="2.1" fill="' + SKIN + '" ' + OUT + "/>";
    };
    if (rank === 13) {
      if (suit === "D") {
        // The King of Diamonds carries an axe.
        return (
          '<path d="M51 45.5V13" stroke="#7a4a1c" stroke-width="1.7" stroke-linecap="round"/>' +
          shape("M51.8 14.5q5.5-1.5 5 5.5-.3 3.5-1.5 4.5-1-3.5-3.5-3.5z", STEEL) +
          hand(51, 41.5)
        );
      }
      return (
        shape("M51.5 38V14.5l1.1-3 1.1 3V38z", STEEL) +
        '<path d="M52.6 15v22" ' + FINE + ' opacity="0.45"/>' +
        shape("M48.8 37.5h7.6v1.7h-7.6z", GOLD) +
        shape("M51.6 39.2h2v6h-2z", "#7a4a1c") +
        hand(52.6, 42)
      );
    }
    if (rank === 12) {
      // A flower in the colours of the house.
      var petal = suit === "H" || suit === "D" ? RED : pal.trim === RED ? RED : "#fff";
      return (
        '<path d="M19 45.5q-1.6-9 0-18" fill="none" stroke="#0a7a2a" stroke-width="1.1"/>' +
        shape("M18.4 36q-4-.5-4.5-4 3.5.3 4.5 4zM19.3 33.5q3.7-.8 4.2-4-3.3.5-4.2 4z", "#12a03a", 'stroke="#0a5a20" stroke-width="0.35"') +
        '<g fill="' + petal + '" stroke="' + INK + '" stroke-width="0.4"><circle cx="19" cy="22.3" r="1.9"/><circle cx="21.6" cy="24.2" r="1.9"/><circle cx="20.6" cy="27.2" r="1.9"/><circle cx="17.4" cy="27.2" r="1.9"/><circle cx="16.4" cy="24.2" r="1.9"/></g>' +
        '<circle cx="19" cy="25.1" r="1.5" fill="' + GOLD + '" stroke="' + INK + '" stroke-width="0.4"/>' +
        hand(19, 41.5)
      );
    }
    // The Jack carries a pike.
    return (
      '<path d="M53.2 45.5V17" stroke="#7a4a1c" stroke-width="1.5" stroke-linecap="round"/>' +
      shape("M53.2 11l2.2 5-2.2 2.5-2.2-2.5z", STEEL) +
      shape("M50.3 18.3h5.8v1.5h-5.8z", pal.trim, 'stroke="' + INK + '" stroke-width="0.4"') +
      hand(53.2, 41.5)
    );
  }

  // Which way each figure looks: 1 to the right, -1 to the left.
  function courtFacing(card) {
    var right = card.suit === "H" || card.suit === "S";
    if (card.rank === 12) right = !right;
    return right ? 1 : -1;
  }

  function courtHalf(card) {
    var pal = COURT[card.suit];
    var headY = card.rank === 13 ? 24.8 : 24;
    return (
      courtBody(card.rank, pal) +
      courtProp(card.rank, card.suit, pal) +
      '<g transform="translate(35.5 ' + headY + ')">' +
      courtHead(card.rank, pal, courtFacing(card)) +
      "</g>" +
      pip(card.suit, 17.5, 14.7, 7.5, false, isRed(card.suit) ? RED : "#000")
    );
  }

  var faceCache = {};

  function faceSVG(card) {
    var k = key(card);
    if (faceCache[k]) return faceCache[k];
    var fill = isRed(card.suit) ? RED : "#000";
    var label = RANKS[card.rank];
    var index =
      '<text x="7.5" y="12.5" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-weight="bold" font-size="' +
      (label === "10" ? 10 : 12) +
      '"' +
      (label === "10" ? ' letter-spacing="-1"' : "") +
      ' fill="' +
      fill +
      '">' +
      label +
      "</text>" +
      pip(card.suit, 7.5, 20, 9, false, fill);
    var out = [
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 71 96" width="71" height="96">',
      '<rect x="0.5" y="0.5" width="70" height="95" rx="3" fill="#fff" stroke="#000"/>',
      index,
      '<g transform="rotate(180 35.5 48)">' + index + "</g>",
    ];
    if (card.rank === 1) {
      out.push(pip(card.suit, C, 48, card.suit === "S" ? 38 : 24, false, fill));
    } else if (card.rank <= 10) {
      PIPS[card.rank].forEach(function (p) {
        out.push(pip(card.suit, p[0], p[1], 14, p[1] > 48, fill));
      });
    } else {
      var half = courtHalf(card);
      out.push('<rect x="12.5" y="9.5" width="46" height="77" fill="#fff"/>');
      out.push('<svg x="12.5" y="9.5" width="46" height="38.5" viewBox="12.5 9.5 46 38.5">' + half + "</svg>");
      out.push(
        '<g transform="rotate(180 35.5 48)"><svg x="12.5" y="9.5" width="46" height="38.5" viewBox="12.5 9.5 46 38.5">' +
          half +
          "</svg></g>",
      );
      out.push('<rect x="12.5" y="9.5" width="46" height="77" fill="none" stroke="#000"/>');
      out.push('<path d="M12.5 48h46" stroke="#000" stroke-width="0.6"/>');
    }
    out.push("</svg>");
    faceCache[k] = out.join("");
    return faceCache[k];
  }

  // ---------------------------------------------------------------- backs

  var BACKS = [
    { id: "weave", name: "Blue weave" },
    { id: "weave-red", name: "Red weave" },
    { id: "fish", name: "Fish" },
    { id: "castle", name: "Castle" },
    { id: "roses", name: "Roses" },
    { id: "shell", name: "Shells" },
  ];

  var BACK_ART = {
    fish:
      '<svg viewBox="0 0 63 88" preserveAspectRatio="none" aria-hidden="true">' +
      '<path d="M0 22q8-4 16 0t16 0 16 0 16 0M0 50q8-4 16 0t16 0 16 0 16 0M0 76q8-4 16 0t16 0 16 0 16 0" fill="none" stroke="#60a0ff" stroke-width="1.5"/>' +
      '<path d="M10 36q10-10 24 0-14 10-24 0z" fill="#ffb000" stroke="#000" stroke-width=".8"/><path d="M34 36l7-5v10z" fill="#ffb000" stroke="#000" stroke-width=".8"/><circle cx="15" cy="34.5" r="1.3"/>' +
      '<path d="M50 62q-9-8-20 0 11 8 20 0z" fill="#ff6060" stroke="#000" stroke-width=".8"/><path d="M30 62l-6-4v8z" fill="#ff6060" stroke="#000" stroke-width=".8"/><circle cx="45" cy="60.6" r="1.2"/>' +
      '<circle cx="44" cy="22" r="2" fill="none" stroke="#fff"/><circle cx="47" cy="15" r="1.4" fill="none" stroke="#fff"/></svg>',
    castle:
      '<svg viewBox="0 0 63 88" preserveAspectRatio="none" aria-hidden="true">' +
      '<circle cx="48" cy="16" r="7" fill="#ffffc0"/>' +
      '<path d="M8 88V44h5v-5h4v5h5v-5h4v5h5V30h4v-5h4v5h4v-5h4v5h4v14h5v-5h4v5h5v44z" fill="#909090" stroke="#000" stroke-width=".8"/>' +
      '<path d="M27 88V70q4-6 9 0v18zM16 56h3v5h-3zM44 56h3v5h-3zM34 38h3v5h-3z" fill="#000"/>' +
      '<path d="M0 88h63v-6H0z" fill="#008000"/></svg>',
    roses:
      '<svg viewBox="0 0 63 88" preserveAspectRatio="none" aria-hidden="true">' +
      '<g fill="#e00000" stroke="#600" stroke-width=".6">' +
      '<circle cx="16" cy="18" r="6"/><circle cx="47" cy="18" r="6"/><circle cx="31.5" cy="44" r="8"/><circle cx="16" cy="70" r="6"/><circle cx="47" cy="70" r="6"/></g>' +
      '<g fill="none" stroke="#600" stroke-width=".7"><path d="M13 18q3-4 6 0M44 18q3-4 6 0M27 44q4-6 9 0M13 70q3-4 6 0M44 70q3-4 6 0"/></g>' +
      '<g fill="#008000"><path d="M22 22l6 4-7 1zM41 22l-6 4 7 1zM22 66l6-4-7-1zM41 66l-6-4 7-1z"/></g></svg>',
    shell:
      '<svg viewBox="0 0 63 88" preserveAspectRatio="none" aria-hidden="true">' +
      '<g fill="#ffd0a0" stroke="#804000" stroke-width=".8">' +
      '<path d="M8 30q12-24 24 0z"/><path d="M31 58q12-24 24 0z"/><path d="M8 84q12-24 24 0z"/></g>' +
      '<g stroke="#804000" stroke-width=".6"><path d="M20 30l-8-7M20 30l-3-10M20 30l3-10M20 30l8-7M43 58l-8-7M43 58l-3-10M43 58l3-10M43 58l8-7M20 84l-8-7M20 84l-3-10M20 84l3-10M20 84l8-7"/></g></svg>',
  };

  function backHTML(design) {
    return '<div class="cb cb-' + design + '">' + (BACK_ART[design] || "") + "</div>";
  }

  // ---------------------------------------------------------------- markup

  /**
   * One positioned card. `opts`: x, y, z, faceDown, back (design id),
   * attrs (extra attribute string), cls (extra class names).
   */
  function cardHTML(card, opts) {
    var down = opts.faceDown;
    return (
      '<div class="card' +
      (down ? " down" : "") +
      (opts.cls ? " " + opts.cls : "") +
      '" style="left:' +
      opts.x +
      "px;top:" +
      opts.y +
      "px;z-index:" +
      (opts.z || 1) +
      '"' +
      (opts.attrs || "") +
      (down ? "" : ' aria-label="' + name(card) + '"') +
      ">" +
      (down ? backHTML(opts.back || "weave") : faceSVG(card)) +
      "</div>"
    );
  }

  /** An empty place a card can go: a pile's outline on the felt. */
  function slotHTML(x, y, attrs, cls, inner) {
    return (
      '<div class="slot' +
      (cls ? " " + cls : "") +
      '" style="left:' +
      x +
      "px;top:" +
      y +
      'px"' +
      (attrs || "") +
      ">" +
      (inner || "") +
      "</div>"
    );
  }

  var imageCache = {};
  /** A face as an <img>, for drawing on a canvas (Solitaire's win). */
  function image(card) {
    var k = key(card);
    if (!imageCache[k]) {
      var img = new Image();
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(faceSVG(card));
      imageCache[k] = img;
    }
    return imageCache[k];
  }

  // ---------------------------------------------------------------- layout

  /**
   * Fit a card table to its window. Below `minW` x `minH` the table is drawn
   * at that size and scaled down, so a narrow window still shows every pile.
   */
  function fitBoard(host, board, minW, minH) {
    var w = Math.max(1, host.clientWidth);
    var h = Math.max(1, host.clientHeight);
    var scale = Math.min(1, w / minW, h / minH);
    var bw = w / scale;
    var bh = h / scale;
    board.style.width = bw + "px";
    board.style.height = bh + "px";
    board.style.transform = scale < 1 ? "scale(" + scale + ")" : "";
    return { scale: scale, width: bw, height: bh };
  }

  function overlap(a, b) {
    var w = Math.min(a.right, b.right) - Math.max(a.left, b.left);
    var h = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
    return w > 0 && h > 0 ? w * h : 0;
  }

  /** The legal drop target under a dragged card: the one it overlaps most. */
  function bestTarget(board, rect, isLegal) {
    var best = null;
    var bestArea = 0;
    board.querySelectorAll("[data-target]").forEach(function (el) {
      var id = el.getAttribute("data-target");
      var area = overlap(rect, el.getBoundingClientRect());
      if (area > bestArea && isLegal(id)) {
        best = id;
        bestArea = area;
      }
    });
    return best;
  }

  /**
   * Pointer handling for a card table.
   *
   *   pick(el)           -> { els, payload } to start dragging, or null
   *   drop(payload, r)   -> true if the cards moved (the game redraws), false
   *                         to send them back
   *   click(el), dblclick(el), rightclick(e)
   *
   * `el` is the element under the pointer that carries data-hit, or null for
   * bare felt. Double-click is detected here rather than left to the browser,
   * so a double tap works on a touch screen too.
   */
  function dragController(board, opts) {
    var press = null;
    var last = { key: null, time: 0 };

    function scale() {
      return opts.scale ? opts.scale() : 1;
    }

    function onDown(e) {
      if (e.button !== 0 || press) return;
      var target = e.target.closest("[data-hit]");
      press = { x: e.clientX, y: e.clientY, target: target, dragging: false, refused: false };
      if (e.pointerType === "mouse") e.preventDefault();
      document.addEventListener("pointermove", onMove);
      document.addEventListener("pointerup", onUp);
      document.addEventListener("pointercancel", onCancel);
    }

    function startDrag() {
      var picked = press.target ? opts.pick(press.target) : null;
      if (!picked || !picked.els.length) {
        press.refused = true;
        return;
      }
      press.dragging = true;
      press.els = picked.els;
      press.payload = picked.payload;
      press.orig = picked.els.map(function (el) {
        return { left: parseFloat(el.style.left), top: parseFloat(el.style.top), z: el.style.zIndex };
      });
      picked.els.forEach(function (el, i) {
        el.classList.add("lifted");
        el.style.zIndex = String(1000 + i);
      });
    }

    function onMove(e) {
      if (!press) return;
      if (!press.dragging) {
        if (press.refused) return;
        if (Math.abs(e.clientX - press.x) + Math.abs(e.clientY - press.y) < 5) return;
        startDrag();
        if (!press.dragging) return;
      }
      var s = scale();
      var dx = (e.clientX - press.x) / s;
      var dy = (e.clientY - press.y) / s;
      press.els.forEach(function (el, i) {
        el.style.left = press.orig[i].left + dx + "px";
        el.style.top = press.orig[i].top + dy + "px";
      });
    }

    function finish() {
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onCancel);
      var p = press;
      press = null;
      return p;
    }

    function sendBack(p) {
      p.els.forEach(function (el, i) {
        el.classList.add("returning");
        el.style.left = p.orig[i].left + "px";
        el.style.top = p.orig[i].top + "px";
      });
      setTimeout(function () {
        p.els.forEach(function (el, i) {
          el.classList.remove("returning", "lifted");
          el.style.zIndex = p.orig[i].z;
        });
      }, 160);
    }

    function onUp(e) {
      var p = finish();
      if (!p) return;
      if (p.dragging) {
        var moved = opts.drop(p.payload, p.els[0].getBoundingClientRect(), e);
        if (!moved) sendBack(p);
        return;
      }
      if (p.refused && Math.abs(e.clientX - p.x) + Math.abs(e.clientY - p.y) >= 5) return;
      var k = p.target ? p.target.getAttribute("data-hit") : "felt";
      var now = Date.now();
      if (last.key === k && now - last.time < 450) {
        last = { key: null, time: 0 };
        if (opts.dblclick) opts.dblclick(p.target, e);
      } else {
        last = { key: k, time: now };
        if (opts.click) opts.click(p.target, e);
      }
    }

    function onCancel() {
      var p = finish();
      if (p && p.dragging) sendBack(p);
    }

    board.addEventListener("pointerdown", onDown);
    board.addEventListener("contextmenu", function (e) {
      e.preventDefault();
      if (opts.rightclick) opts.rightclick(e);
    });

    return {
      cancel: function () {
        var p = finish();
        if (p && p.dragging) sendBack(p);
      },
      busy: function () {
        return !!(press && press.dragging);
      },
    };
  }

  Win95.Cards = {
    W: W,
    H: H,
    SUITS: SUITS,
    RANKS: RANKS,
    BACKS: BACKS,
    isRed: isRed,
    newDeck: newDeck,
    shuffle: shuffle,
    name: name,
    key: key,
    faceSVG: faceSVG,
    backHTML: backHTML,
    cardHTML: cardHTML,
    slotHTML: slotHTML,
    image: image,
    fitBoard: fitBoard,
    bestTarget: bestTarget,
    dragController: dragController,
  };
})();
