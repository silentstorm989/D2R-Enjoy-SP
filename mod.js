if (D2RMM.getVersion == null || D2RMM.getVersion() < 1.6) {
  D2RMM.error('Requires D2RMM version 1.6 or higher.');
  return;
}

// ExpandedInventory
{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    const id = row.class;
    const classes = [
      'Amazon',
      'Assassin',
      'Barbarian',
      'Druid',
      'Necromancer',
      'Paladin',
      'Sorceress',
    ];
    if (
      classes.indexOf(id) !== -1 ||
      classes.map((cls) => `${cls}2`).indexOf(id) !== -1
    ) {
      row.gridX = 13;
      row.gridY = 8;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const profileHDFilename = 'global\\ui\\layouts\\_profilehd.json';
  const profileHD = D2RMM.readJson(profileHDFilename);
  profileHD.RightPanelRect_ExpandedInventory = {
    x: -1394 - (1382 - 1162),
    y: -651,
    width: 1382,
    height: 1507,
  };
  profileHD.PanelClickCatcherRect_ExpandedInventory = {
    x: 0,
    y: 0,
    width: 1162,
    height: 1507,
  };
  // offset the right hinge so that it doesn't overlap with content of the right panel
  profileHD.RightHingeRect = { x: 1076 + 20, y: 630 };
  profileHD.RightHingeRect_ExpandedInventory = {
    x: 1076 + (1382 - 1162) + 20,
    y: 630,
  };
  D2RMM.writeJson(profileHDFilename, profileHD);

  const profileLVFilename = 'global\\ui\\layouts\\_profilelv.json';
  const profileLV = D2RMM.readJson(profileLVFilename);
  profileLV.RightPanelRect_ExpandedInventory = {
    x: -1346 - (1382 - 1162) * 1.16,
    y: -856,
    width: 1382,
    height: 1507,
    scale: 1.16,
  };
  D2RMM.writeJson(profileLVFilename, profileLV);

  const playerInventoryOriginalLayoutFilename =
    'global\\ui\\layouts\\playerinventoryoriginallayout.json';
  const playerInventoryOriginalLayout = D2RMM.readJson(
    playerInventoryOriginalLayoutFilename
  );
  // TODO: new sprite & layout for classic UI
  playerInventoryOriginalLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 13;
      child.fields.cellCount.y = 8;
    }
  });
  D2RMM.writeJson(
    playerInventoryOriginalLayoutFilename,
    playerInventoryOriginalLayout
  );

  const playerInventoryOriginalLayoutHDFilename =
    'global\\ui\\layouts\\playerinventoryoriginallayouthd.json';
  const playerInventoryOriginalLayoutHD = D2RMM.readJson(
    playerInventoryOriginalLayoutHDFilename
  );
  playerInventoryOriginalLayoutHD.fields.rect =
    '$RightPanelRect_ExpandedInventory';
  playerInventoryOriginalLayoutHD.children =
    playerInventoryOriginalLayoutHD.children.filter((child) => {
      if (child.name === 'background') {
        child.fields.filename = 'PANEL\\Inventory\\Classic_Background_Expanded';
      }
      if (child.name === 'click_catcher') {
        child.fields.rect = { x: 0, y: 45, width: 1093, height: 1495 };
      }
      if (child.name === 'RightHinge') {
        child.fields.rect = '$RightHingeRect_ExpandedInventory';
      }
      if (child.name === 'title') {
        child.fields.rect = {
          x: 91 + (1382 - 1162) / 2,
          y: 64,
          width: 972,
          height: 71,
        };
      }
      if (child.name === 'close') {
        child.fields.rect.x = child.fields.rect.x + (1382 - 1162);
      }
      if (child.name === 'grid') {
        child.fields.cellCount.x = 13;
        child.fields.cellCount.y = 8;
        child.fields.rect.x = child.fields.rect.x - 37;
        child.fields.rect.y = child.fields.rect.y - 229;
      }
      if (child.name === 'slot_right_arm') {
        child.fields.rect.x = child.fields.rect.x - 14;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_left_arm') {
        child.fields.rect.x = child.fields.rect.x + 227;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_torso') {
        child.fields.rect.x = child.fields.rect.x + 101;
        child.fields.rect.y = child.fields.rect.y - 229;
      }
      if (child.name === 'slot_head') {
        child.fields.rect.x = child.fields.rect.x - 144;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_gloves') {
        child.fields.rect.x = child.fields.rect.x + 231;
        child.fields.rect.y = child.fields.rect.y - 233;
      }
      if (child.name === 'slot_feet') {
        child.fields.rect.x = child.fields.rect.x - 26;
        child.fields.rect.y = child.fields.rect.y - 231;
      }
      if (child.name === 'slot_belt') {
        child.fields.rect.x = child.fields.rect.x + 101;
        child.fields.rect.y = child.fields.rect.y - 234;
      }
      if (child.name === 'slot_neck') {
        child.fields.rect.x = child.fields.rect.x + 99;
        child.fields.rect.y = child.fields.rect.y - 182;
      }
      if (child.name === 'slot_right_hand') {
        child.fields.rect.x = child.fields.rect.x + 474;
        child.fields.rect.y = child.fields.rect.y - 466;
      }
      if (child.name === 'slot_left_hand') {
        child.fields.rect.x = child.fields.rect.x + 232;
        child.fields.rect.y = child.fields.rect.y - 466;
      }
      if (child.name === 'gold_amount' || child.name === 'gold_button') {
        child.fields.rect.x = child.fields.rect.x - 291;
        child.fields.rect.y = child.fields.rect.y - 1267;
      }
      return true;
    });
  D2RMM.writeJson(
    playerInventoryOriginalLayoutHDFilename,
    playerInventoryOriginalLayoutHD
  );

  const playerInventoryExpansionLayoutHDFilename =
    'global\\ui\\layouts\\playerinventoryexpansionlayouthd.json';
  const playerInventoryExpansionLayoutHD = D2RMM.readJson(
    playerInventoryExpansionLayoutHDFilename
  );
  playerInventoryExpansionLayoutHD.children =
    playerInventoryExpansionLayoutHD.children.filter((child) => {
      if (child.name === 'click_catcher') {
        // make click catcher work the same way as in the originallayouthd file
        return false;
      }
      if (child.name === 'background') {
        child.fields.filename = 'PANEL\\Inventory\\Background_Expanded';
      }
      if (
        child.name === 'background_right_arm' ||
        child.name === 'background_right_arm_selected' ||
        child.name === 'weaponswap_right_arm' ||
        child.name === 'text_i_left' ||
        child.name === 'text_ii_left'
      ) {
        child.fields.rect.x = child.fields.rect.x - 14;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (
        child.name === 'background_left_arm' ||
        child.name === 'background_left_arm_selected' ||
        child.name === 'weaponswap_left_arm' ||
        child.name === 'text_i_right' ||
        child.name === 'text_ii_right'
      ) {
        child.fields.rect.x = child.fields.rect.x + 227;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      return true;
    });
  D2RMM.writeJson(
    playerInventoryExpansionLayoutHDFilename,
    playerInventoryExpansionLayoutHD
  );

  const playerInventoryOriginalControllerLayoutHDFilename =
    'global\\ui\\layouts\\controller\\playerinventoryoriginallayouthd.json';
  const playerInventoryOriginalControllerLayoutHD = D2RMM.readJson(
    playerInventoryOriginalControllerLayoutHDFilename
  );
  playerInventoryOriginalControllerLayoutHD.children.forEach((child) => {
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/InventoryPanel/V2/InventoryBG_Classic_Expanded';
      child.fields.rect.x = child.fields.rect.x - 166;
      child.fields.rect.y = child.fields.rect.y - 160;
    }
    if (child.name === 'grid') {
      child.fields.rect.x = child.fields.rect.x - 132;
      child.fields.rect.y = child.fields.rect.y - 344;
    }
    if (child.name === 'slot_right_arm') {
      child.fields.rect.x = child.fields.rect.x - 99;
      child.fields.rect.y = child.fields.rect.y - 60;
    }
    if (child.name === 'slot_left_arm') {
      child.fields.rect.x = child.fields.rect.x + 123;
      child.fields.rect.y = child.fields.rect.y - 62;
    }
    if (child.name === 'slot_torso') {
      child.fields.rect.x = child.fields.rect.x + 6;
      child.fields.rect.y = child.fields.rect.y - 199;
    }
    if (child.name === 'slot_head') {
      child.fields.rect.x = child.fields.rect.x - 239;
      child.fields.rect.y = child.fields.rect.y + 21;
    }
    if (child.name === 'slot_gloves') {
      child.fields.rect.x = child.fields.rect.x + 146;
      child.fields.rect.y = child.fields.rect.y - 282;
    }
    if (child.name === 'slot_feet') {
      child.fields.rect.x = child.fields.rect.x - 130;
      child.fields.rect.y = child.fields.rect.y - 281;
    }
    if (child.name === 'slot_belt') {
      child.fields.rect.x = child.fields.rect.x + 7;
      child.fields.rect.y = child.fields.rect.y - 185;
    }
    if (child.name === 'slot_neck') {
      child.fields.rect.x = child.fields.rect.x - 3;
      child.fields.rect.y = child.fields.rect.y - 167;
    }
    if (child.name === 'slot_right_hand') {
      child.fields.rect.x = child.fields.rect.x + 389;
      child.fields.rect.y = child.fields.rect.y - 417;
    }
    if (child.name === 'slot_left_hand') {
      child.fields.rect.x = child.fields.rect.x + 126;
      child.fields.rect.y = child.fields.rect.y - 417;
    }
    if (child.name === 'Belt') {
      child.fields.rect.x = child.fields.rect.x + 15;
      child.fields.rect.y = child.fields.rect.y + 595;
    }
    if (child.name === 'gold_amount' || child.name === 'gold_button') {
      child.fields.rect.x = child.fields.rect.x - 464;
      child.fields.rect.y = child.fields.rect.y + 20;
    }
  });
  D2RMM.writeJson(
    playerInventoryOriginalControllerLayoutHDFilename,
    playerInventoryOriginalControllerLayoutHD
  );

  const playerInventoryExpansionControllerLayoutHDFilename =
    'global\\ui\\layouts\\controller\\playerinventoryexpansionlayouthd.json';
  const playerInventoryExpansionControllerLayoutHD = D2RMM.readJson(
    playerInventoryExpansionControllerLayoutHDFilename
  );
  playerInventoryExpansionControllerLayoutHD.children.forEach((child) => {
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/InventoryPanel/V2/InventoryBG_Expanded';
    }
    if (
      child.name === 'background_right_arm' ||
      child.name === 'background_right_arm_selected' ||
      child.name === 'WeaponSwapRightLegend' ||
      child.name === 'text_i_left' ||
      child.name === 'text_ii_left'
    ) {
      child.fields.rect.x = child.fields.rect.x - 99;
      child.fields.rect.y = child.fields.rect.y - 60;
    }
    if (
      child.name === 'background_left_arm' ||
      child.name === 'background_left_arm_selected' ||
      child.name === 'WeaponSwapLeftLegend' ||
      child.name === 'text_i_right' ||
      child.name === 'text_ii_right'
    ) {
      child.fields.rect.x = child.fields.rect.x + 123;
      child.fields.rect.y = child.fields.rect.y - 62;
    }
  });
  D2RMM.writeJson(
    playerInventoryExpansionControllerLayoutHDFilename,
    playerInventoryExpansionControllerLayoutHD
  );
}

// ExpandedCube
{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    if (
      row.class === 'Transmogrify Box Page 1' ||
      row.class === 'Transmogrify Box2'
    ) {
      row.gridX = 6;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const profileHDFilename = 'global\\ui\\layouts\\_profilehd.json';
  const profileHD = D2RMM.readJson(profileHDFilename);
  D2RMM.writeJson(profileHDFilename, profileHD);

  const profileLVFilename = 'global\\ui\\layouts\\_profilelv.json';
  const profileLV = D2RMM.readJson(profileLVFilename);
  D2RMM.writeJson(profileLVFilename, profileLV);

  const horadricCubeLayoutFilename =
    'global\\ui\\layouts\\horadriccubelayout.json';
  const horadricCubeLayout = D2RMM.readJson(horadricCubeLayoutFilename);
  horadricCubeLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 6;
      child.fields.cellCount.y = 4;
      // TODO: shift left for new sprite
      // one cell = 29px, add 1 column to left and 2 columns to right of original space
      child.fields.rect.x = child.fields.rect.x - 29;
    }
    // TODO: new sprite
  });
  D2RMM.writeJson(horadricCubeLayoutFilename, horadricCubeLayout);

  const horadricCubeLayoutHDFilename =
    'global\\ui\\layouts\\horadriccubelayouthd.json';
  const horadricCubeHDLayout = D2RMM.readJson(horadricCubeLayoutHDFilename);
  horadricCubeHDLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 6;
      child.fields.cellCount.y = 4;
      child.fields.rect.x = child.fields.rect.x - 144;
    }
    if (child.name === 'background') {
      child.fields.filename = 'PANEL\\Horadric_Cube\\HoradricCube_BG_Expanded';
    }
  });
  D2RMM.writeJson(horadricCubeLayoutHDFilename, horadricCubeHDLayout);

  const controllerHoradricCubeHDLayoutFilename =
    'global\\ui\\layouts\\controller\\horadriccubelayouthd.json';
  const controllerHoradricCubeHDLayout = D2RMM.readJson(
    controllerHoradricCubeHDLayoutFilename
  );
  controllerHoradricCubeHDLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.rect.x = child.fields.rect.x - 142;
    }
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/HoradricCube/V2/HoradricCubeBG_Expanded';
    }
  });
  D2RMM.writeJson(
    controllerHoradricCubeHDLayoutFilename,
    controllerHoradricCubeHDLayout
  );
}

// MercEquip
{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    if (row.class === 'Hireling') {
      row.rArmLeft = 340;
      row.rArmRight = 395;
      row.rArmTop = 27;
      row.rArmBottom = 139;
      row.torsoLeft = 453;
      row.torsoRight = 509;
      row.lArmLeft = 566;
      row.lArmRight = 621;
      row.lArmTop = 27;
      row.lArmBottom = 139;
      row.headLeft = 455;
      row.headRight = 509;
      row.neckLeft = 529;
      row.neckRight = 552;
      row.neckTop = 35;
      row.neckBottom = 59;
      row.neckWidth = 23;
      row.neckHeight = 24;
      row.rHandLeft = 415;
      row.rHandRight = 438;
      row.rHandTop = 180;
      row.rHandBottom = 204;
      row.rHandWidth = 23;
      row.rHandHeight = 24;
      row.lHandLeft = 529;
      row.lHandRight = 552;
      row.lHandTop = 180;
      row.lHandBottom = 204;
      row.lHandWidth = 23;
      row.lHandHeight = 24;
      row.beltLeft = 456;
      row.beltRight = 508;
      row.beltTop = 179;
      row.beltBottom = 204;
      row.beltWidth = 52;
      row.beltHeight = 25;
      row.feetLeft = 572;
      row.feetRight = 626;
      row.feetTop = 155;
      row.feetBottom = 207;
      row.feetWidth = 54;
      row.feetHeight = 52;
      row.glovesLeft = 341;
      row.glovesRight = 395;
      row.glovesTop = 153;
      row.glovesBottom = 205;
      row.glovesWidth = 54;
      row["glovesHeight\r"] = 53;
    } else if (row.class === 'Hireling2') {
      row.rArmLeft = 420;
      row.rArmRight = 475;
      row.rArmTop = 89;
      row.rArmBottom = 201;
      row.torsoLeft = 533;
      row.torsoRight = 589;
      row.lArmLeft = 646;
      row.lArmRight = 701;
      row.lArmTop = 89;
      row.lArmBottom = 201;
      row.headLeft = 535;
      row.headRight = 589;
      row.neckLeft = 609;
      row.neckRight = 632;
      row.neckTop = 95;
      row.neckBottom = 119;
      row.neckWidth = 23;
      row.neckHeight = 24;
      row.rHandLeft = 495;
      row.rHandRight = 518;
      row.rHandTop = 240;
      row.rHandBottom = 264;
      row.rHandWidth = 23;
      row.rHandHeight = 24;
      row.lHandLeft = 609;
      row.lHandRight = 632;
      row.lHandTop = 240;
      row.lHandBottom = 264;
      row.lHandWidth = 23;
      row.lHandHeight = 24;
      row.beltLeft = 536;
      row.beltRight = 588;
      row.beltTop = 239;
      row.beltBottom = 264;
      row.beltWidth = 52;
      row.beltHeight = 25;
      row.feetLeft = 652;
      row.feetRight = 706;
      row.feetTop = 217;
      row.feetBottom = 269;
      row.feetWidth = 54;
      row.feetHeight = 52;
      row.glovesLeft = 421;
      row.glovesRight = 475;
      row.glovesTop = 215;
      row.glovesBottom = 268;
      row.glovesWidth = 54;
      row["glovesHeight\r"] = 53;
    } else if (['Amazon', 'Sorceress', 'Necromancer', 'Paladin', 'Barbarian', 'Druid', 'Assassin']
      .indexOf(row.class) !== -1) {
      row.gridTop = 211;
      row.gridBottom = 441;
      row.rArmTop = 27;
      row.rArmBottom = 139;
      row.lArmLeft = 566;
      row.lArmRight = 621;
      row.lArmTop = 27;
      row.lArmBottom = 139;
      row.feetTop = 155;
      row.feetBottom = 207;
      row.glovesTop = 153;
      row.glovesBottom = 205;
    } else if (['Amazon2', 'Sorceress2', 'Necromancer2', 'Paladin2', 'Barbarian2', 'Druid2', 'Assassin2']
      .indexOf(row.class) !== -1) {
      row.gridTop = 271;
      row.gridBottom = 501;
      row.rArmTop = 89;
      row.rArmBottom = 201;
      row.lArmLeft = 646;
      row.lArmRight = 701;
      row.lArmTop = 89;
      row.lArmBottom = 201;
      row.feetTop = 217;
      row.feetBottom = 269;
      row.glovesTop = 215;
      row.glovesBottom = 268;
    } else if (row.class === 'Transmogrify Box Page 1') {
      row.gridLeft = 16;
      row.gridRight = 303;
      row.gridTop = 16;
      row.gridBottom = 254;
    } else if (row.class === 'Transmogrify Box2') {
      row.gridLeft = 97;
      row.gridRight = 382;
      row.gridTop = 75;
      row.gridBottom = 304;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const itemTypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemTypesFilename);
  itemtypes.rows.forEach((row) => {
    if (['Ring', 'Amulet', 'Boots', 'Gloves', 'Belt'].indexOf(row.ItemType) !== -1) {
      row.Equiv2 = 'merc';
    } else if (row.ItemType === 'Helm') {
      row.Code = 'merc';
      row.Equiv1 = '';
    } else if (row.ItemType === 'Expansion') {
      row["*eol\r"] = '0';
    }
  });
  const itemRecipe = {
    ItemType: 'Merc Equip',
    Code: 'helm',
    Equiv1: 'armo',
    Equiv2: 'merc',
    Repair: 1,
    Body: 1,
    BodyLoc1: 'head',
    BodyLoc2: 'head',
    Throwable: 0,
    Reload: 0,
    ReEquip: 0,
    AutoStack: 0,
    Rare: 1,
    Normal: 0,
    Beltable: 0,
    MaxSockets1: 2,
    MaxSocketsLevelThreshold1: 25,
    MaxSockets2: 2,
    MaxSocketsLevelThreshold2: 40,
    MaxSockets3: 3,
    TreasureClass: 0,
    Rarity: 3,
    VarInvGfx: 0,
    StorePage: 'armo',
    '*eol\r': 0,
  };
  itemtypes.rows.push({
    ...itemRecipe
  });
  D2RMM.writeTsv(itemTypesFilename, itemtypes);

  const hirelingHDFilename = 'global\\ui\\layouts\\hirelinginventorypanelhd.json';
  const hirelingHD = D2RMM.readJson(hirelingHDFilename);
  hirelingHD.children.forEach((child) => {
    if (child.type === 'ClickCatcherWidget') {
      child.fields.rect = {
        x: -180,
        y: -200,
        width: 1362,
        height: 1727
      };
    } else if (child.type === 'TextBoxWidget') {
      if (child.name === 'Title') {
        child.fields.rect = {
          x: 481,
          y: -69,
          width: 196,
          height: 196
        };
      } else if (child.name === 'CharacterName') {
        child.fields.rect.x = 121;
        child.fields.rect.y = 849;
        child.fields.style.alignment.v = 'bottom';
      } else if (child.name === 'HPTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HPStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HireTypeText') {
        child.fields.rect.x = 126;
        child.fields.rect.y = 937;
        child.fields.rect.width = 100;
        child.fields.rect.height = 30;
        child.fields.style.pointSize = '$XMediumFontSize';
        delete child.fields.style.alignment.v;
      } else if (child.name === 'StrengthTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'StrengthStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClassTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClass') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      }
    } else if (child.type === 'InventorySlotWidget') {
      if (child.name === 'slot_head') {
        child.fields.rect.y = 113;
        delete child.fields['gemSocketFilename'];
      } else if (child.name === 'slot_torso') {
        child.fields.rect.x = 483;
        child.fields.rect.y = 348;
        delete child.fields['gemSocketFilename'];
      } else if (child.name === 'slot_right_arm') {
        child.fields.rect.x = 110;
        child.fields.rect.y = 156;
        child.fields.location = 'right_arm';
      } else if (child.name === 'slot_left_arm') {
        child.fields.rect.x = 863;
        child.fields.rect.y = 156;
        child.fields.location = 'left_arm';
      }
    } else if (child.type === 'ExpBarWidget') {
      child.fields.rect.y = 913;
    } else if (child.type === 'ButtonWidget') {
      if (child.name === 'AdvancedStats') {
        child.fields.rect.x = 1104;
        child.fields.rect.y = 713;
        delete child.fields['hoveredFrame'];
        delete child.fields['pressLabelOffset'];
      } else if (child.name === 'CloseButton') {
        delete child.fields['sound'];
      }
    } else if (child.type === 'HirelingSkillIconWidget') {
      if (child.name === 'Skill0') {
        child.fields.rect.x = 673;
      } else if (child.name === 'Skill1') {
        child.fields.rect.x = 780;
      } else if (child.name === 'Skill2') {
        child.fields.rect.x = 887;
      }
      child.fields.rect.y = 1007;
      child.fields.rect.scale = 0.60;
    } else if (child.type === 'Widget') {
      if (child.name === 'Damage') {
        child.fields.rect.y = 1269;
      }
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_belt",
    fields: {
      rect: {
        x: 483,
        y: 688,
        width: 196,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "belt",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Belt",
      isHireable: true
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_right_hand",
    fields: {
      rect: {
        x: 720,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "right_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_left_hand",
    fields: {
      rect: {
        x: 344,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "left_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_gloves",
    fields: {
      rect: {
        x: 108,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "gloves",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Glove",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_feet",
    fields: {
      rect: {
        x: 861,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "feet",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Boots",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_neck",
    fields: {
      rect: {
        x: 720,
        y: 271,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "neck",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Amulet",
      isHireable: false
    }
  });
  D2RMM.writeJson(hirelingHDFilename, hirelingHD);

  const profileCHDFilename = 'global\\ui\\layouts\\controller\\_profilehd.json';
  const profileCHD = D2RMM.readJson(profileCHDFilename);
  profileCHD.ConsoleLeftPanelAnchor = { "x": 0.521, "y": 0.387 };
  D2RMM.writeJson(profileCHDFilename, profileCHD);

  const hirelingCHDFilename = 'global\\ui\\layouts\\controller\\hirelinginventorypanelhd.json';
  const hirelingCHD = D2RMM.readJson(hirelingCHDFilename);
  delete hirelingCHD['basedOn'];
  hirelingCHD.fields = {
    rect: "$ConsoleLeftPanelRect",
    anchor: "$ConsoleLeftPanelAnchor",
    bowBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Bow",
    spearBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Spear",
    longswordBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_LongSword",
    shieldBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Shield",
    twoHandSwordBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_2HSword"
  };
  hirelingCHD.children.forEach((child) => {
    if (child.type === 'ImageWidget') {
      child.fields = {
        rect: {
          x: 0,
          y: 0
        },
        filename: "\\PANEL\\Hireling\\HirelingPanel"
      }
    } else if (child.type === 'ClickCatcherWidget') {
      child.fields = {
        rect: {
          x: -180,
          y: -200,
          width: 1362,
          height: 1727
        }
      };
    } else if (child.type === 'InventorySlotWidget') {
      if (child.name === 'slot_head') {
        child.fields.rect.x = 481;
        child.fields.rect.y = 113;
        child.fields.location = 'head';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_HeadArmor';
      } else if (child.name === 'slot_torso') {
        child.fields.rect.x = 483;
        child.fields.rect.y = 348;
        child.fields.location = 'torso';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_ChestArmor';
      } else if (child.name === 'slot_right_arm') {
        child.fields.rect.x = 110;
        child.fields.rect.y = 156;
        child.fields.location = 'right_arm';
        child.fields.gemSocketFilename = 'PANEL\\gemsocket';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_Weapon';
      } else if (child.name === 'slot_left_arm') {
        child.fields.rect.x = 863;
        child.fields.rect.y = 156;
        child.fields.location = 'left_arm';
        child.fields.gemSocketFilename = 'PANEL\\gemsocket';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_Weapon';
      }
      child.fields.cellSize = '$ItemCellSize';
      child.fields.isHireable = true;
      delete child.fields['navigation'];
    } else if (child.type === 'TextBoxWidget') {
      if (child.name === 'CharacterName') {
        child.fields.rect.x = 121;
        child.fields.rect.y = 849;
        child.fields.rect.width = 500;
        child.fields.rect.height = 50;
        child.fields.style.fontColor = '$FontColorWhite';
        child.fields.style.pointSize = '$LargeFontSize';
        child.fields.style.dropShadow = '$DefaultDropShadow';
      } else if (child.name === 'HPTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrlif';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HPStat') {
        child.fields.rect.y = 1020;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HireTypeText') {
        child.fields.rect.x = 126;
        child.fields.rect.y = 937;
        child.fields.rect.width = 100;
        child.fields.rect.height = 30;
        child.fields.text = '@strchrlvl';
        child.fields.style.fontColor = '$FontColorWhite';
        child.fields.style.pointSize = '$XMediumFontSize';
        child.fields.style.dropShadow = '$DefaultDropShadow';
        delete child.fields.style.alignment.v;
      } else if (child.name === 'StrengthTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrstr';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'StrengthStat') {
        child.fields.rect.y = 1102;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrdex';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexStat') {
        child.fields.rect.y = 1186;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrskm';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageStat') {
        child.fields.rect.y = 1269;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClassTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrdef';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClass') {
        child.fields.rect.y = 1353;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'FireText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'ColdText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'LightningText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'PoisonText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      }
    } else if (child.type === 'ExpBarWidget') {
      child.fields = {
        rect: {
          x: 139,
          y: 913,
          width: 888,
          height: 10
        },
        filename: "PANEL\\Hireling\\Hireling_ExpBar",
        frame: 0,
        isHireling: true,
        hitMargin: {
          top: 15,
          bottom: 15
        }
      }
    } else if (child.type === 'HirelingSkillIconWidget') {
      if (child.name === 'Skill0') {
        child.fields.rect.x = 673;
      } else if (child.name === 'Skill1') {
        child.fields.rect.x = 780;
      } else if (child.name === 'Skill2') {
        child.fields.rect.x = 887;
      }

      child.fields.rect.y = 1007;
      child.fields.rect.scale = 0.60;
    } else if (child.type === 'Widget') {
      if (child.name === 'Damage') {
        child.fields = {
          rect: {
            x: 328,
            y: 1269,
            width: 237,
            height: 59
          }
        };
        child.children.forEach((wdChild) => {
          wdChild.fields.rect.width = 237;
          wdChild.fields.rect.height = 59;
          wdChild.fields.style.pointSize = '$SmallPanelFontSize';
          if (wdChild.name === 'DamageStatTop') {
            wdChild.fields.rect.y = 0;
          } else if (wdChild.name === 'DamageStatBottom') {
            wdChild.fields.rect.y = -1;
          }
        });
      }
    }
  });
  hirelingCHD.children.push({
    type: "ImageWidget",
    name: "LeftHinge",
    fields: {
      rect: "$LeftHingeRect",
      filename: "$LeftHingeSprite"
    }
  });
  hirelingCHD.children.push({
    type: "TextBoxWidget",
    name: "Title",
    fields: {
      rect: {
        x: 481,
        y: -69,
        width: 196,
        height: 196
      },
      style: "$StyleTitleBlock",
      text: "@MiniPanelHireinv"
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_belt",
    fields: {
      rect: {
        x: 483,
        y: 688,
        width: 196,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "belt",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Belt",
      isHireable: true
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_right_hand",
    fields: {
      rect: {
        x: 720,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "right_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_left_hand",
    fields: {
      rect: {
        x: 344,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "left_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_gloves",
    fields: {
      rect: {
        x: 108,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "gloves",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Glove",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_feet",
    fields: {
      rect: {
        x: 861,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "feet",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Boots",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_neck",
    fields: {
      rect: {
        x: 720,
        y: 271,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "neck",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Amulet",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "ButtonWidget",
    name: "AdvancedStats",
    fields: {
      rect: {
        x: 1104,
        y: 713
      },
      filename: "PANEL\\Character_Sheet\\AdvancedStatsButton",
      onClickMessage: "HirelingInventoryPanelMessage:ToggleAdvancedStats",
      tooltipString: "@AdvancedStats"
    }
  });
  hirelingCHD.children.push({
    type: "ButtonWidget",
    name: "CloseButton",
    fields: {
      rect: {
        x: 1075,
        y: 9
      },
      filename: "PANEL\\closebtn_4x",
      hoveredFrame: 3,
      onClickMessage: "HirelingInventoryPanelMessage:Close",
      tooltipString: "@strClose"
    }
  });
  D2RMM.writeJson(hirelingCHDFilename, hirelingCHD);
}

// Improved rune drop and general drop rate
{
  function SetDefaultQuestProp(row, setNoDrop) {
    row.Unique = 1024;
    row.Set = 800;
    row.Rare = 800;
    row.Magic = 800;
    if (setNoDrop) {
      row.NoDrop = 0;
    }
  }
  function SetQuestProp(row, pOne) {
    SetDefaultQuestProp(row, true);
    row.Item1 = 'gld,mul=2000';
    row.Prob1 = pOne;
  }
  const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
  const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
  treasureclassex.rows.forEach((row) => {
    const treasureClass = row['Treasure Class'];
    // Improved rune drop
    {
      if (treasureClass === 'Runes 17') {
        row.Prob1 = 3;
        row.Prob2 = 16;
      }
      if (treasureClass === 'Runes 16') {
        row.Prob3 = 15;
        row.Prob4 = 7;
      }
      if (treasureClass === 'Runes 15') {
        row.Prob3 = 14;
        row.Prob4 = 7;
      }
      if (treasureClass === 'Runes 14') {
        row.Prob3 = 13;
        row.Prob4 = 8;
      }
      if (treasureClass === 'Runes 13') {
        row.Prob3 = 12;
        row.Prob4 = 8;
      }
      if (treasureClass === 'Runes 12') {
        row.Prob3 = 11;
        row.Prob4 = 4;
      }
      if (treasureClass === 'Runes 11') {
        row.Prob3 = 10;
      }
      if (treasureClass === 'Runes 10') {
        row.Prob3 = 9;
      }
      if (treasureClass === 'Runes 9') {
        row.Prob3 = 8;
      }
      if (treasureClass === 'Runes 8') {
        row.Prob3 = 7;
      }
      if (treasureClass === 'Runes 7') {
        row.Prob3 = 6;
      }
      if (treasureClass === 'Runes 6') {
        row.Prob3 = 5;
      }
      if (treasureClass === 'Runes 5') {
        row.Prob3 = 4;
      }
      if (treasureClass === 'Runes 4') {
        row.Prob3 = 3;
      }
      if (treasureClass === 'Runes 3') {
        row.Prob3 = 3;
      }
      if (treasureClass === 'Runes 2') {
        row.Prob3 = 3;
      }
    }

    // Increased general drop rate from Bosses
    {
      if (treasureClass === 'Andariel') {
        SetQuestProp(row, 5);
        row.Prob2 = 15;
        row.Item3 = 'Act 2 Equip C';
        row.Prob3 = 5;
        row.Item4 = 'rin';
        row.Prob4 = 2;
      }
      if (treasureClass === 'Andariel (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 5;
        row.Prob2 = 19;
        row.Item3 = 'Act 2 (N) Equip C';
        row.Prob3 = 6;
        row.Item4 = 'Act 2 (N) Good';
        row.Prob4 = 3;
      }
      if (treasureClass === 'Andariel (H)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=4048';
        row.Prob1 = 5;
        row.Item3 = 'Act 2 (H) Equip C';
        row.Prob3 = 7;
        row.Prob4 = 5;
        row.Prob5 = 0;
      }
      if (treasureClass === 'Andarielq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Item2 = 'Act 2 Equip C';
        row.Prob2 = 5;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Andarielq (N)') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Item2 = 'Act 2 (N) Good';
        row.Prob2 = 1;
        row.Item3 = 'Act 1 (N) Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Andarielq (H)') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 1 (H) Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Baal') {
        SetQuestProp(row, 0);
        row.Item2 = 'Act 1 (N) Equip C';
        row.Item3 = 'Act 1 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Baal (N)') {
        SetQuestProp(row, 0);
        row.Item1 = 'gld,mul=3536';
        row.Item2 = 'Act 1 (H) Equip C';
        row.Item3 = 'Act 1 (H) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Baal (H)') {
        SetQuestProp(row, 0);
        row.Item1 = 'gld,mul=4048';
        row.Item2 = 'Act 5 (H) Equip C';
        row.Item3 = 'Act 5 (H) Good';
        row.Prob3 = 3;
        row.Item4 = 'fed';
        row.Prob4 = 0;
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Baalq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 25;
      }
      if (treasureClass === 'Baalq (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'Act 1 (H) Equip C';
        row.Prob1 = 26;
        row.Prob2 = 5;
      }
      if (treasureClass === 'Baalq (H)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'Act 5 (H) Equip C';
        row.Prob1 = 26;
        row.Prob2 = 5;
        row.Prob3 = 0;
        row.Item4 = 'uar'
        row.Prob4 = 1;
      }
      if (treasureClass === 'Blood Raven') {
        SetQuestProp(row, 6);
        row.Item2 = 'Act 1 Equip A';
        row.Prob2 = 14;
        row.Item3 = 'Act 1 Good';
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Blood Raven (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 6;
        row.Item2 = 'Act 1 (N) Equip A';
        row.Prob2 = 14;
        row.Item3 = 'Act 1 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Blood Raven (H)') {
        row.Unique = 1024;
        row.Set = 850;
        row.Rare = 850;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=4048';
        row.Prob1 = 11;
        row.Item2 = 'Act 1 (H) Equip A';
        row.Prob2 = 33;
        row.Item3 = '6lw';
        row.Prob3 = 1;
        row.Prob4 = 9;
      }
      if (treasureClass === 'Countess') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Countess (N)') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 (N) Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Countess (H)') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 (H) Good';
        row.Prob3 = 1;
        row.Item4 = 'pk1';
        row.Prob4 = 1;
      }
      if (treasureClass === 'Countess Rune') {
        row.Item1 = 'Runes 5';
      }
      if (treasureClass === 'Countess Rune (N)') {
        row.Item1 = 'Runes 11';
        row.Prob1 = 3;
        row.Item2 = 'Runes 8';
        row.Prob2 = 1;
      }
      if (treasureClass === 'Countess Rune (H)') {
        row.Item1 = 'Runes 11';
        row.Prob1 = 6;
        row.Item2 = 'Runes 15';
        row.Prob2 = 2;
        row.Item3 = 'Runes 16';
        row.Prob3 = 1;
        row.Item4 = 'Runes 13';
        row.Prob4 = 6;
        row.Item5 = 'Runes 14';
        row.Prob5 = 2;
      }
      if (treasureClass === 'Council') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=3280';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Council (N)') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=4536';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Council (H)') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=5048';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 (H) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Cow') {
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Cow King') {
        row.Picks = 7;
      }
      if (treasureClass === 'Diablo') {
        SetQuestProp(row, 4);
        row.Prob2 = 15;
        row.Item3 = 'Act 5 Equip C';
        row.Prob3 = 3;
        row.Prob4 = 5;
      }
      if (treasureClass === 'Diabloq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 5 Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Duriel') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel (N)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel (H)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel - Base') {
        SetQuestProp(row, 11);
        row.Item3 = 'Act 3 Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
      }
      if (treasureClass === 'Duriel (N) - Base') {
        SetQuestProp(row, 11);
        row.Item1 = 'gld,mul=3536';
        row.Item3 = 'Act 3 (N) Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
      }
      if (treasureClass === 'Duriel (H) - Base') {
        SetQuestProp(row, 11);
        row.Item1 = 'gld,mul=4048';
        row.Item3 = 'Act 3 (H) Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
        row.Prob5 = 0;
      }
      if (treasureClass === 'Durielq') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq (N)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq (H)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 3 Equip B';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Durielq (N) - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 3 (N) Equip B';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Durielq (H) - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 44;
        row.Prob2 = 3;
        row.Item3 = 'Act 3 (H) Equip B';
        row.Prob3 = 6;
        row.Item4 = 'xrs';
        row.Prob4 = 1;
      }
      if (treasureClass === 'Flying Scimitar') {
        row.Unique = 999;
        row.Set = 899;
        row.Rare = 850;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Prob1 = 11;
        row.Prob2 = 7;
      }
      if (treasureClass === 'Griswold') {
        row.Picks = 4;
        SetDefaultQuestProp(row, false);
        row.Item1 = 'Act 2 Uitem C';
        row.Prob1 = 8;
        row.Item2 = 'Act 2 Melee A';
        row.Prob2 = 15;
        row.Item3 = 'bsd';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Griswold (N)') {
        row.Picks = 4;
        row.Unique = 999;
        row.Set = 999;
        row.Rare = 800;
        row.Magic = 800;
        row.Item1 = 'Act 1 (H) Uitem C';
        row.Prob1 = 8;
        row.Item2 = 'Act 1 (N) Melee B';
        row.Prob2 = 15;
      }
      if (treasureClass === 'Haphesto') {
        row.Picks = 4;
        SetQuestProp(row, 5);
        row.Item2 = 'Act 4 Equip A';
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 4;
        row.Prob4 = 2;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Izual') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 8;
        row.Item3 = 'Act 4 Good';
        row.Prob3 = 8;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Mephisto') {
        SetQuestProp(row, 4);
        row.Prob2 = 15;
        row.Item3 = 'Act 5 Equip A';
        row.Prob3 = 3;
        row.Prob4 = 5;
      }
      if (treasureClass === 'Mephistoq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 5 Equip A';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Nihlathak') {
        row.Picks = 6;
        SetQuestProp(row, 5);
        row.Item2 = 'Act 5 Equip C';
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Radament') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 3;
        row.Item2 = 'Act 3 Equip A';
        row.Prob2 = 15;
        row.Item3 = 'Act 3 Good';
        row.Prob3 = 7;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Radament (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 3;
        row.Item2 = 'Act 3 (N) Equip A';
        row.Prob2 = 3;
        row.Item3 = 'Act 3 (N) Good';
        row.Prob3 = 15;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Smith') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Smith (N)') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Smith (H)') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Summoner') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 4;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 2;
        row.Item4 = 'Act 3 Equip A';
        row.Prob4 = 4;
        row.Item5 = '';
        row.Prob5 = '';
      }
    }

    // Increased more drop rates.
    {
      const chestLevel = ["A", "B", "C"];
      const diffLevel = ['', '(N) ', '(H) '];
      for (let acts = 1; acts <= 5; acts = acts + 1) {
        for (let level = 0; level < 3; level = level + 1) {
          for (let diff = 0; diff < 3; diff = diff + 1) {
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.level = 70;
                  }
                  if (chestLevel[level] === "C") {
                    row.level = 70;
                    row.Prob1 = 0;
                    row.Prob2 = 0;
                    row.Prob3 = 0;
                    row.Prob4 = 0;
                    row.Prob6 = 5;
                    row.Prob7 = 5;
                    row.Prob8 = 5;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Cast ${chestLevel[level]}`) {
              // Act 1
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                // Normal
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic ${chestLevel[level]}`;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              // Act 2
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                else if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                else {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
              }
              // Act 3
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
              }
              else if (acts == 4) {
                if (chestLevel[level] === "A") {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                }
                else {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                }
              }
              else {
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
                else {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Prob3 = 2;
              row.Prob4 = 7;
              row.Item5 = '';
              row.Prob5 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Chest ${chestLevel[level]}`) {
              // Act 1
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  }
                  else {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip B`;
                  }
                  else {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  }
                  row.Prob2 = 10;
                }
                // Normal
                else {
                  if (chestLevel[level] === "C") {
                    row.Prob2 = 15;
                  }
                  else {
                    row.Prob2 = 10;
                  }
                  row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                }
              }
              // Act 2
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 2 (H) Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 2 (N) Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 Equip A`;
                    row.Prob2 = 10;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 2 Equip B`;
                    row.Prob2 = 10;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 2 Equip C`;
                    row.Prob2 = 15;
                  }
                }
              }
              // Act 3
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 3 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 3 (N) Equip B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 Equip A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 3 Equip B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 3 Equip C`;
                  }
                }
                row.Prob2 = 10;
              }
              // Act 4
              else if (acts === 4) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 4 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 (N) Equip A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 4 (N) Equip B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 4 (N) Equip C`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 Equip A`;
                  }
                  else {
                    row.Item2 = `Act 4 Equip B`;
                  }
                }
                row.Prob2 = 10;
              }
              // Act
              else if (acts === 5) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 (N) Equip B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 Equip B`;
                  }
                }
                row.Prob2 = 10;
              }
              else {
                row.Prob2 = 15;
              }
              row.NoDrop = 50;
              row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              row.Prob3 = 4;
              row.Item4 = '';
              row.Prob4 = '';
              row.Item5 = '';
              row.Prob5 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.Prob1 = 1;
                    row.Prob3 = 4;
                    row.Prob4 = 6;
                    row.Prob5 = 12;
                    row.Prob6 = 6;
                    row.Prob7 = 4;
                    row.Prob8 = 3;
                  }
                  if (chestLevel[level] === "C") {
                    row.Prob1 = 1;
                    row.Prob3 = 4;
                    row.Prob4 = 6;
                    row.Prob5 = 11;
                    row.Prob6 = 5;
                    row.Prob7 = 5;
                    row.Prob8 = 5;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}H2H ${chestLevel[level]}`) {
              if (acts === 2) {
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 Good`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                }
                else {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              else {
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Prob3 = 2;
              row.Item4 = '';
              row.Prob4 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Melee ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.Prob1 = 2;
                    row.Prob3 = 6;
                    row.Prob4 = 3;
                    row.Prob5 = 14;
                    row.Prob6 = 7;
                    row.Prob7 = 3;
                    row.Prob8 = 3;
                  }
                  if (chestLevel[level] === "C") {
                    row.Prob1 = 1;
                    row.Prob3 = 3;
                    row.Prob4 = 3;
                    row.Prob5 = 11;
                    row.Prob6 = 7;
                    row.Prob7 = 4;
                    row.Prob8 = 4;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Miss ${chestLevel[level]}`) {
              row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              row.Prob3 = 2;
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow A`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
              }
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else if (acts === 5) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else {
                row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
              }
              row.Prob4 = 6;
              row.Item5 = 'Ammo';
              row.Prob5 = 3;
              row.Item6 = '';
              row.Prob6 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super ${chestLevel[level]}`) {
              // Super A
              if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super A`) {
                // Act 1
                if (acts === 1) {
                  // Normal
                  if (diff === 0) {
                    row.level = 5;
                    row.Prob1 = 14;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 14;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 2
                if (acts === 2) {
                  // Normal
                  if (diff === 0) {
                    row.level = 12;
                    row.Item2 = `Act 1 Equip C`;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item2 = `Act 1 (N) Equip C`;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 1 (H) Equip C`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 3
                if (acts === 3) {
                  // Normal
                  if (diff === 0) {
                    row.Item2 = `Act 2 Equip C`;
                    row.Prob1 = 15;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item2 = `Act 2 (N) Equip C`;
                    row.Prob1 = 15;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 2 (H) Equip C`;
                    row.Prob1 = 15;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 4
                if (acts === 4) {
                  // Normal
                  if (diff === 0) {
                    row.level = 26;
                    row.Prob1 = 12;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 12;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  row.Prob2 = 12;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
                // Act 5
                if (acts === 5) {
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  else if (diff === 0) {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  else {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
              }
              // Super B
              else if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super B`) {
                if (acts === 1) {
                  // Normal
                  if (diff === 0) {
                    row.level = 7;
                    row.Prob1 = 14;
                    row.Prob3 = 6;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 14;
                    row.Prob3 = 6;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                    row.Prob3 = 6;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 2) {
                  row.Prob1 = 15;
                  row.Prob3 = 6;
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 3) {
                  // Hell
                  if (diff === 2) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  else if (diff === 1) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  else {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  row.Prob1 = 15;
                  row.Prob3 = 6;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 4) {
                  if (diff === 1) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob1 = 12;
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                    row.Prob3 = 3;
                  }
                  else if (diff === 2) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob1 = 12;
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                    row.Prob3 = 3;
                  }
                  else {
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                    row.Prob1 = 12;
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob3 = 3;
                  }
                  row.Prob2 = 12;
                }
                if (acts === 5) {
                  row.Prob1 = 15;
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Prob3 = 3;
                }
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;

              }
              // Super C
              else if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super C`) {
                // Act 1
                if (acts === 1) {
                  // Hell
                  if (diff === 2) {
                    row.Prob2 = 2;
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 2 ${diffLevel[diff]}Equip A`;
                  row.Prob3 = 6;
                }
                // Act 2
                if (acts === 2) {
                  // Hell
                  if (diff === 2) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  // Normal
                  if (diff === 0) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 3 ${diffLevel[diff]}Equip A`;
                  row.Prob3 = 6;
                } // Act 3
                if (acts === 3) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip A`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 4 ${diffLevel[diff]}Equip A`;
                  row.Prob2 = 2;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 4
                if (acts === 4) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 12;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  row.Prob2 = 12;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
                // Act 5
                if (acts === 5) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob2 = 3;
                }
              }
              row.Picks = 5;
              row.Set = 1024;
              row.Rare = 800;
              row.Magic = 800;
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super ${chestLevel[level]}x`) {
              row.group = 18;
              row.Picks = 5;
              row.Magic = 900;
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Unique ${chestLevel[level]}`) {
              if (acts === 2) {
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 3 ${diffLevel[diff]}Good`;
                  }
                  else {
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                }
                else {
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              else {
                row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Picks = 3;
              row.Unique = 933;
              row.Set = 1010;
              row.Rare = 800;
              row.Magic = 800;
              row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
              row.Prob1 = 11;
              row.Prob2 = 1;
            }
          }
        }
      }
    }

    if (treasureClass === 'Chipped Gem') {
      row.Prob1 = 10;
      row.Prob2 = 6;
      row.Prob3 = 6;
      row.Prob4 = 6;
      row.Prob5 = 6;
      row.Prob6 = 6;
      row.Prob7 = 6;
    }
    if (treasureClass === 'Flawed Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Normal Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Flawless Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Perfect Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Runes 12') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 13') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 14') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 15') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 16') {
      row.Item4 = 'Runes 2';
    }
  });
  D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation Armor.txt
{
  const NameAndRarity = [
    { name: "Cap", rarity: 1 },
    { name: "Skull Cap", rarity: 1 },
    { name: "Helm", rarity: 1 },
    { name: "Full Helm", rarity: 1 },
    { name: "Great Helm", rarity: 2 },
    { name: "Crown", rarity: 2 },
    { name: "Mask", rarity: 2 },
    { name: "Quilted Armor", rarity: 1 },
    { name: "Leather Armor", rarity: 2 },
    { name: "Hard Leather Armor", rarity: 1 },
    { name: "Studded Leather ", rarity: 1 },
    { name: "Ring Mail", rarity: 1 },
    { name: "Scale Mail", rarity: 1 },
    { name: "Chain Mail", rarity: 1 },
    { name: "Breast Plate", rarity: 1 },
    { name: "Splint Mail", rarity: 1 },
    { name: "Plate Mail", rarity: 1 },
    { name: "Field Plate", rarity: 1 },
    { name: "Gothic Plate", rarity: 1 },
    { name: "Full Plate Mail", rarity: 1 },
    { name: "Ancient Armor", rarity: 1 },
    { name: "Light Plate", rarity: 1 },
    { name: "Buckler", rarity: 1 },
    { name: "Small Shield", rarity: 1 },
    { name: "Large Shield", rarity: 1 },
    { name: "Kite Shield", rarity: 2 },
    { name: "Tower Shield", rarity: 2 },
    { name: "Gothic Shield", rarity: 2 },
    { name: "Leather Gloves", rarity: 1 },
    { name: "Heavy Gloves", rarity: 1 },
    { name: "Chain Gloves", rarity: 2 },
    { name: "Light Gauntlets", rarity: 3 },
    { name: "Gauntlets", rarity: 4 },
    { name: "Boots", rarity: 1 },
    { name: "Heavy Boots", rarity: 1 },
    { name: "Chain Boots", rarity: 2 },
    { name: "Light Plated Boots", rarity: 3 },
    { name: "Greaves", rarity: 2 },
    { name: "Sash", rarity: 1 },
    { name: "Light Belt", rarity: 1 },
    { name: "Belt", rarity: 2 },
    { name: "Heavy Belt", rarity: 2 },
    { name: "Plated Belt", rarity: 3 },
    { name: "Bone Helm", rarity: 2 },
    { name: "Bone Shield", rarity: 2 },
    { name: "Spiked Shield", rarity: 3 },
    { name: "War Hat", rarity: 1 },
    { name: "Sallet", rarity: 2 },
    { name: "Casque", rarity: 2 },
    { name: "Basinet", rarity: 2 },
    { name: "Winged Helm", rarity: 2 },
    { name: "Grand Crown", rarity: 2 },
    { name: "Death Mask", rarity: 2 },
    { name: "Ghost Armor", rarity: 1 },
    { name: "Serpentskin Armor", rarity: 2 },
    { name: "Demonhide Armor", rarity: 3 },
    { name: "Trellised Armor", rarity: 2 },
    { name: "Linked Mail", rarity: 2 },
    { name: "Tigulated Mail", rarity: 2 },
    { name: "Mesh Armor", rarity: 2 },
    { name: "Cuirass", rarity: 2 },
    { name: "Russet Armor", rarity: 2 },
    { name: "Templar Coat", rarity: 2 },
    { name: "Sharktooth Armor", rarity: 2 },
    { name: "Embossed Plate", rarity: 2 },
    { name: "Chaos Armor", rarity: 3 },
    { name: "Ornate Armor", rarity: 2 },
    { name: "Mage Plate", rarity: 2 },
    { name: "Defender", rarity: 2 },
    { name: "Round Shield", rarity: 3 },
    { name: "Scutum", rarity: 2 },
    { name: "Dragon Shield", rarity: 2 },
    { name: "Pavise", rarity: 2 },
    { name: "Ancient Shield", rarity: 2 },
    { name: "Demonhide Gloves", rarity: 1 },
    { name: "Sharkskin Gloves", rarity: 1 },
    { name: "Heavy Bracers", rarity: 2 },
    { name: "Battle Gauntlets", rarity: 3 },
    { name: "War Gauntlets", rarity: 2 },
    { name: "Demonhide Boots", rarity: 1 },
    { name: "Sharkskin Boots", rarity: 1 },
    { name: "Mesh Boots", rarity: 2 },
    { name: "Battle Boots", rarity: 3 },
    { name: "War Boots", rarity: 2 },
    { name: "Demonhide Sash", rarity: 1 },
    { name: "Sharkskin Belt", rarity: 1 },
    { name: "Mesh Belt", rarity: 2 },
    { name: "Battle Belt", rarity: 2 },
    { name: "War Belt", rarity: 2 },
    { name: "Grim Helm", rarity: 2 },
    { name: "Grim Shield", rarity: 2 },
    { name: "Barbed Shield", rarity: 2 },
    { name: "Wolf Head", rarity: 1 },
    { name: "Hawk Helm", rarity: 1 },
    { name: "Antlers", rarity: 1 },
    { name: "Falcon Mask", rarity: 1 },
    { name: "Spirit Mask", rarity: 1 },
    { name: "Jawbone Cap", rarity: 2 },
    { name: "Fanged Helm", rarity: 2 },
    { name: "Horned Helm", rarity: 2 },
    { name: "Assault Helmet", rarity: 2 },
    { name: "Avenger Guard", rarity: 2 },
    { name: "Targe", rarity: 2 },
    { name: "Rondache", rarity: 2 },
    { name: "Heraldic Shield", rarity: 2 },
    { name: "Aerin Shield", rarity: 2 },
    { name: "Crown Shield", rarity: 2 },
    { name: "Preserved Head", rarity: 1 },
    { name: "Zombie Head", rarity: 1 },
    { name: "Unraveller Head", rarity: 1 },
    { name: "Gargoyle Head", rarity: 1 },
    { name: "Demon Head", rarity: 1 },
    { name: "Circlet", rarity: 1 },
    { name: "Coronet", rarity: 1 },
    { name: "Tiara", rarity: 1 },
    { name: "Diadem", rarity: 1 },
    { name: "Shako", rarity: 1 },
    { name: "Hydraskull", rarity: 4 },
    { name: "Armet", rarity: 4 },
    { name: "Giant Conch", rarity: 4 },
    { name: "Spired Helm", rarity: 4 },
    { name: "Corona", rarity: 4 },
    { name: "Demonhead", rarity: 4 },
    { name: "Dusk Shroud", rarity: 1 },
    { name: "Wyrmhide", rarity: 2 },
    { name: "Scarab Husk", rarity: 2 },
    { name: "Wire Fleece", rarity: 2 },
    { name: "Diamond Mail", rarity: 2 },
    { name: "Loricated Mail", rarity: 2 },
    { name: "Boneweave", rarity: 2 },
    { name: "Great Hauberk", rarity: 2 },
    { name: "Balrog Skin", rarity: 2 },
    { name: "Hellforge Plate", rarity: 2 },
    { name: "Kraken Shell", rarity: 2 },
    { name: "Lacquered Plate", rarity: 2 },
    { name: "Shadow Plate", rarity: 2 },
    { name: "Sacred Armor", rarity: 3 },
    { name: "Archon Plate", rarity: 3 },
    { name: "Heater", rarity: 2 },
    { name: "Luna", rarity: 3 },
    { name: "Hyperion", rarity: 2 },
    { name: "Monarch", rarity: 3 },
    { name: "Aegis", rarity: 2 },
    { name: "Ward", rarity: 2 },
    { name: "Bramble Mitts", rarity: 1 },
    { name: "Vampirebone Gloves", rarity: 1 },
    { name: "Vambraces", rarity: 2 },
    { name: "Crusader Gauntlets", rarity: 2 },
    { name: "Ogre Gauntlets", rarity: 2 },
    { name: "Wyrmhide Boots", rarity: 1 },
    { name: "Scarabshell Boots", rarity: 1 },
    { name: "Boneweave Boots", rarity: 2 },
    { name: "Mirrored Boots", rarity: 2 },
    { name: "Myrmidon Greaves", rarity: 2 },
    { name: "Spiderweb Sash", rarity: 1 },
    { name: "Vampirefang Belt", rarity: 1 },
    { name: "Mithril Coil", rarity: 2 },
    { name: "Troll Belt", rarity: 2 },
    { name: "Colossus Girdle", rarity: 2 },
    { name: "Bone Visage", rarity: 2 },
    { name: "Troll Nest", rarity: 2 },
    { name: "Blade Barrier", rarity: 2 },
    { name: "Alpha Helm", rarity: 1 },
    { name: "Griffon Headress", rarity: 1 },
    { name: "Hunter's Guise", rarity: 1 },
    { name: "Sacred Feathers", rarity: 1 },
    { name: "Totemic Mask", rarity: 1 },
    { name: "Jawbone Visor", rarity: 2 },
    { name: "Lion Helm", rarity: 2 },
    { name: "Rage Mask", rarity: 2 },
    { name: "Savage Helmet", rarity: 2 },
    { name: "Slayer Guard", rarity: 2 },
    { name: "Akaran Targe", rarity: 2 },
    { name: "Akaran Rondache", rarity: 2 },
    { name: "Protector Shield", rarity: 2 },
    { name: "Gilded Shield", rarity: 2 },
    { name: "Royal Shield", rarity: 2 },
    { name: "Mummified Trophy", rarity: 1 },
    { name: "Fetish Trophy", rarity: 1 },
    { name: "Sexton Trophy", rarity: 1 },
    { name: "Cantor Trophy", rarity: 1 },
    { name: "Heirophant Trophy", rarity: 1 },
    { name: "Blood Spirit", rarity: 1 },
    { name: "Sun Spirit", rarity: 1 },
    { name: "Earth Spirit", rarity: 1 },
    { name: "Sky Spirit", rarity: 1 },
    { name: "Dream Spirit", rarity: 1 },
    { name: "Carnage Helm", rarity: 2 },
    { name: "Fury Visor", rarity: 2 },
    { name: "Destroyer Helm", rarity: 2 },
    { name: "Conqueror Crown", rarity: 2 },
    { name: "Guardian Crown", rarity: 2 },
    { name: "Sacred Targe", rarity: 2 },
    { name: "Sacred Rondache", rarity: 2 },
    { name: "Ancient Shield", rarity: 2 },
    { name: "Zakarum Shield", rarity: 2 },
    { name: "Vortex Shield", rarity: 2 },
    { name: "Minion Skull", rarity: 1 },
    { name: "Hellspawn Skull", rarity: 1 },
    { name: "Overseer Skull", rarity: 1 },
    { name: "Succubus Skull", rarity: 1 },
    { name: "Bloodlord Skull", rarity: 1 },
  ];

  function findRarityByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.rarity : null; // Return null if not found
  }

  const armorFilename = 'global\\excel\\armor.txt';
  const armor = D2RMM.readTsv(armorFilename);
  armor.rows.forEach((row) => {
    const theName = row['name'];
    const rarity = findRarityByName(theName);
    if (rarity !== null) {
      row.rarity = rarity;
    }
  });
  D2RMM.writeTsv(armorFilename, armor);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation cubemain.txt'
// Removed the gems from the rune upgrade recipes. 
// - Up until Pul: 3 of the same runes = next rune
// - After Pul: 2 of a kind + jewel = next rune
// Removed Hel-Rune from the de-socked recipe. New recipe: 1 Town Scroll + Socketed Item (destroys gems).
{
  const runeChangesExtra = [
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Thul Runes + 1 Chipped Topaz -> Amn Rune",
      "new_description": "3 Thul Runes -> Amn Rune",
      "numinputs": "3",
      "input1": "\"r10,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Amn Runes + 1 Chipped Amethyst -> Sol Rune",
      "new_description": "3 Amn Runes -> Sol Rune",
      "numinputs": "3",
      "input1": "\"r11,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Sol Runes + 1 Chipped Sapphire -> Shael Rune",
      "new_description": "3 Sol Runes -> Shael Rune",
      "numinputs": "3",
      "input1": "\"r12,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": ""
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Shael Runes + 1 Chipped Ruby -> Dol Rune",
      "new_description": "3 Shael Runes -> Dol Rune",
      "numinputs": "3",
      "input1": "\"r13,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Dol Runes + 1 Chipped Emerald -> Hel Rune",
      "new_description": "3 Dol Runes -> Hel Rune",
      "numinputs": "3",
      "input1": "\"r14,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Hel Runes + 1 Chipped Diamond -> Io Rune",
      "new_description": "3 Hel Runes -> Io Rune",
      "numinputs": "3",
      "input1": "\"r15,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Io Runes + 1 Flawed Topaz -> Lum Rune",
      "new_description": "3 Io Runes -> Lum Rune",
      "numinputs": "3",
      "input1": "\"r16,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Lum Runes + 1 Flawed Amethyst -> Ko Rune",
      "new_description": "3 Lum Runes -> Ko Rune",
      "numinputs": "3",
      "input1": "\"r17,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Ko Runes + 1 Flawed Sapphire -> Fal Rune",
      "new_description": "3 Ko Runes -> Fal Rune",
      "numinputs": "3",
      "input1": "\"r18,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Fal Runes + 1 Flawed Ruby -> Lem Rune",
      "new_description": "3 Fal Runes -> Lem Rune",
      "numinputs": "3",
      "input1": "\"r19,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "3 Lem Runes + 1 Flawed Emerald -> Pul Rune",
      "new_description": "3 Lem Runes -> Pul Rune",
      "numinputs": "3",
      "input1": "\"r20,qty=3\"",
      "input2": "",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Pul Runes + 1 Flawed Diamond -> Um Rune",
      "new_description": "2 Pul Runes + 1 Jewel -> Um Rune",
      "numinputs": "3",
      "input1": "\"r21,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Um Runes + 1 Standard Topaz -> Mal Rune",
      "new_description": "2 Um Runes + 1 Jewel -> Mal Rune",
      "numinputs": "3",
      "input1": "\"r22,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Mal Runes + 1 Standard Amethyst -> Ist Rune",
      "new_description": "2 Mal Runes + 1 Jewel -> Ist Rune",
      "numinputs": "3",
      "input1": "\"r23,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Ist Runes + 1 Standard Sapphire -> Gul Rune",
      "new_description": "2 Ist Runes + 1 Jewel -> Gul Rune",
      "numinputs": "3",
      "input1": "\"r24,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Gul Runes + 1 Standard Ruby -> Vex Rune",
      "new_description": "2 Gul Runes + 1 Jewel -> Vex Rune",
      "numinputs": "3",
      "input1": "\"r25,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Vex Runes + 1 Standard Emerald -> Ohm Rune",
      "new_description": "2 Vex Runes + 1 Jewel -> Ohm Rune",
      "numinputs": "3",
      "input1": "\"r26,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Ohm Runes + 1 Standard Diamond -> Lo Rune",
      "new_description": "2 Ohm Runes + 1 Jewel -> Lo Rune",
      "numinputs": "3",
      "input1": "\"r27,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Lo Runes + 1 Flawless Topaz -> Sur Rune",
      "new_description": "2 Lo Runes +1 Jewel -> Sur Rune",
      "numinputs": "3",
      "input1": "\"r28,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Sur Runes + 1 Flawless Amethyst -> Ber Rune",
      "new_description": "2 Sur Runes + 1 Jewel -> Ber Rune",
      "numinputs": "3",
      "input1": "\"r29,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Ber Runes + 1 Flawless Sapphire -> Jah Rune",
      "new_description": "2 Ber Runes + 1 Jewel -> Jah Rune",
      "numinputs": "3",
      "input1": "\"r30,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Jah Runes + 1 Flawless Ruby -> Cham Rune",
      "new_description": "2 Jah Runes + 1 Jewel -> Cham Rune",
      "numinputs": "3",
      "input1": "\"r31,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "2 Cham Runes + 1 Flawless Emerald -> Zod Rune",
      "new_description": "2 Cham Runes + 1 Jewel -> Zod Rune",
      "numinputs": "3",
      "input1": "\"r32,qty=2\"",
      "input2": "jew",
      "input3": "",
      "input4": "",
    },
    {
      "enabled": 1,
      "version": 100,
      "old_description": "1 Twisted Essence of Suffering + 1 Charged Essence of Hatred + 1 Burning Essence of Terror + 1 Festering Essence of Destruction -> Token of Absolution",
      "new_description": "1 Scroll of Town Portal + Scroll of Identify -> Token of Absolution",
      "numinputs": "2",
      "input1": "tsc",
      "input2": "isc",
      "input3": "",
      "input4": "",
    }];

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  cubemain.rows.forEach((row) => {
    const theName = row['description'];
    const change = runeChangesExtra.find(rc => rc.old_description === theName);
    if (change) {
      row['description'] = change.new_description;
      row['numinputs'] = change.numinputs;
      row['input 1'] = change.input1;
      row['input 2'] = change.input2;
      row['input 3'] = change.input3;
      row['input 4'] = change.input4;
    }
  });
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation cubemain.txt'
// New recipes to add sockets to items via the Horadric Cube
{
  const newRecipes = [
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Unique Armor or Weapon -> Socketed Unique 1",
      "numinputs": "2",
      "input1": "any,uni",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Skull + Set Armor or Weapon -> Socketed Set 1",
      "numinputs": "2",
      "input1": "any,set",
      "input2": "skz,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + 1 Normal Torso Armor -> Socketed Torso Armor 1",
      "numinputs": "2",
      "input1": "tors,nor,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + 1 Normal Torso Armor -> Socketed Torso Armor 2",
      "numinputs": "3",
      "input1": "tors,nor,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + 1 Normal Torso Armor -> Socketed Torso Armor 3",
      "numinputs": "4",
      "input1": "tors,nor,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + 1 Normal Torso Armor -> Socketed Torso Armor 4",
      "numinputs": "5",
      "input1": "tors,nor,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Superior Armor -> Socketed Superior Armor 1",
      "numinputs": "2",
      "input1": "tors,hiq,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Superior Armor -> Socketed Superior Armor 2",
      "numinputs": "3",
      "input1": "tors,hiq,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Superior Armor -> Socketed Superior Armor 3",
      "numinputs": "4",
      "input1": "tors,hiq,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + Superior Armor -> Socketed Superior Armor 4",
      "numinputs": "5",
      "input1": "tors,hiq,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Normal Helm -> Socketed Normal Helm 1",
      "numinputs": "2",
      "input1": "helm,nor,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem  + Normal Helm -> Socketed Normal Helm 2",
      "numinputs": "3",
      "input1": "helm,nor,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Normal Helm -> Socketed Normal Helm 3",
      "numinputs": "4",
      "input1": "helm,nor,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Superior Helm -> Socketed Superior Helm 1",
      "numinputs": "2",
      "input1": "helm,hiq,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Superior Helm -> Socketed Superior Helm 2",
      "numinputs": "3",
      "input1": "helm,hiq,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Superior Helm -> Socketed Superior Helm 3",
      "numinputs": "4",
      "input1": "helm,hiq,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Normal Shield -> Socketed Shield 1",
      "numinputs": "2",
      "input1": "shld,nor,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Normal Shield -> Socketed Shield 2",
      "numinputs": "3",
      "input1": "shld,nor,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Normal Shield -> Socketed Shield 3",
      "numinputs": "4",
      "input1": "shld,nor,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + Normal Shield -> Socketed Shield 4",
      "numinputs": "5",
      "input1": "shld,nor,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Superior Shield -> Socketed Superior Shield 1",
      "numinputs": "2",
      "input1": "shld,hiq,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Superior Shield -> Socketed Superior Shield 2",
      "numinputs": "3",
      "input1": "shld,hiq,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Superior Shield -> Socketed Superior Shield 3",
      "numinputs": "4",
      "input1": "shld,hiq,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + Superior Shield -> Socketed Superior Shield 4",
      "numinputs": "5",
      "input1": "shld,hiq,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Normal Weapon -> Socketed Weapon 1",
      "numinputs": "2",
      "input1": "weap,nor,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Normal Weapon -> Socketed Weapon 2",
      "numinputs": "3",
      "input1": "weap,nor,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Normal Weapon -> Socketed Weapon 3",
      "numinputs": "4",
      "input1": "weap,nor,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + Normal Weapon -> Socketed Weapon 4",
      "numinputs": "5",
      "input1": "weap,nor,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Unique Ring + 2 Perfect Skull + Normal Weapon -> Socketed Weapon 5",
      "numinputs": "3",
      "input1": "weap,nor,nos",
      "input2": "ring,uni",
      "input3": "skz,qty=1",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "5",
      "mod1Max": "5",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Unique Ring + 1 Perfect Skull + Normal Weapon -> Socketed Weapon 6",
      "numinputs": "4",
      "input1": "weap,nor,nos",
      "input2": "ring,uni",
      "input3": "skz,qty=2",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "6",
      "mod1Max": "6",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "1 Perfect Gem + Superior Weapon -> Socketed Superior. Weapon 1",
      "numinputs": "2",
      "input1": "weap,hiq,nos",
      "input2": "gem4,qty=1",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "1",
      "mod1Max": "1",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "2 Perfect Gem + Superior Weapon -> Socketed Superior. Weapon 2",
      "numinputs": "3",
      "input1": "weap,hiq,nos",
      "input2": "gem4,qty=2",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "2",
      "mod1Max": "2",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "3 Perfect Gem + Superior Weapon -> Socketed Superior Weapon 3",
      "numinputs": "4",
      "input1": "weap,hiq,nos",
      "input2": "gem4,qty=3",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "3",
      "mod1Max": "3",
    },
    {
      "enabled": 1,
      "version": 100,
      "description": "4 Perfect Gem + Superior Weapon -> Socketed Superior Weapon 4",
      "numinputs": "5",
      "input1": "weap,hiq,nos",
      "input2": "gem4,qty=4",
      "input3": "",
      "input4": "",
      "useitem": "useitem",
      "mod1": "sock",
      "mod1Min": "4",
      "mod1Max": "4",
    }
  ];

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  newRecipes.forEach((newRecipe) => {
    cubemain.rows.push({
      ...newRecipe,
      'description': newRecipe.description,
      'numinputs': newRecipe.numinputs,
      'input 1': "\"" + newRecipe.input1 + "\"",
      'input 2': newRecipe.input2,
      'input 3': newRecipe.input3,
      'input 4': newRecipe.input4,
      'output': "\"" + newRecipe.useitem + "\"",
      'mod 1': newRecipe.mod1,
      'mod 1 min': newRecipe.mod1Min,
      'mod 1 max': newRecipe.mod1Max,
      '*eol\r': 0,
    });
  })
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation Misc.txt
{
  const miscFilename = 'global\\excel\\misc.txt';
  const misc = D2RMM.readTsv(miscFilename);
  misc.rows.forEach((row) => {
    const theName = row['name'];
    if (theName === 'Tome of Town Portal') {
      row.maxstack = 50;
    }
    if (theName === 'Tome of Identify') {
      row.maxstack = 50;
    }
    if (theName === 'Key of Destruction') {
      row.compactsave = 1;
      row.spawnable = 1;
      row.cost = 6666;
      row.PermStoreItem = 1;
      row.multibuy = 1;
      row.AkaraMin = 1;
      row.AkaraMax = 1;
    }
    if (theName === 'Key of Terror') {
      row.compactsave = 1;
      row.spawnable = 1;
      row.cost = 6666;
      row.PermStoreItem = 1;
      row.multibuy = 1;
      row.AkaraMin = 1;
      row.AkaraMax = 1;
    }
    if (theName === 'Key of Hate') {
      row.compactsave = 1;
      row.spawnable = 1;
      row.cost = 6666;
      row.PermStoreItem = 1;
      row.multibuy = 1;
      row.AkaraMin = 1;
      row.AkaraMax = 1;
    }
  });
  D2RMM.writeTsv(miscFilename, misc);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation DifficultyLevels.txt'
{
  const NameAndGamble = [
    { name: "Normal", gambleRare: 35000, gambleSet: 15000, gambleUnique: 10000, },
    { name: "Nightmare", gambleRare: 35000, gambleSet: 15000, gambleUnique: 10000, },
    { name: "Hell", gambleRare: 35000, gambleSet: 15000, gambleUnique: 10000, },
  ];
  const difficultyFilename = 'global\\excel\\difficultylevels.txt';
  const difficulty = D2RMM.readTsv(difficultyFilename);
  difficulty.rows.forEach((row) => {
    const theNameGamble = NameAndGamble.find(ng => ng.name === row['Name']);
    if (theNameGamble) {
      row.GambleRare = theNameGamble.gambleRare;
      row.GambleSet = theNameGamble.gambleSet;
      row.GambleUnique = theNameGamble.gambleUnique;
    }
  });
  D2RMM.writeTsv(difficultyFilename, difficulty);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation Experience.txt'
{
  const ChangeExp = [
    { level: 70, expRatio: 1024, },
    { level: 71, expRatio: 1024, },
    { level: 72, expRatio: 1024, },
    { level: 73, expRatio: 1024, },
    { level: 74, expRatio: 1024, },
    { level: 75, expRatio: 1024, },
    { level: 76, expRatio: 1024, },
    { level: 77, expRatio: 1024, },
    { level: 78, expRatio: 1024, },
    { level: 79, expRatio: 1024, },
    { level: 80, expRatio: 1024, },
    { level: 81, expRatio: 1024, },
    { level: 82, expRatio: 1024, },
    { level: 83, expRatio: 1024, },
    { level: 84, expRatio: 1024, },
    { level: 85, expRatio: 1024, },
    { level: 86, expRatio: 1024, },
    { level: 87, expRatio: 1024, },
    { level: 88, expRatio: 1024, },
    { level: 89, expRatio: 1024, },
    { level: 90, expRatio: 1024, },
    { level: 91, expRatio: 1024, },
    { level: 92, expRatio: 1024, },
    { level: 93, expRatio: 1024, },
    { level: 94, expRatio: 1024, },
    { level: 95, expRatio: 1024, },
    { level: 96, expRatio: 1024, },
    { level: 97, expRatio: 1024, },
    { level: 98, expRatio: 1024, },
  ];

  function findExpRatioByLevel(level) {
    // Parse the level to an integer
    const parsedLevel = parseInt(level, 10);

    if (isNaN(parsedLevel)) {
      return undefined; // Or handle appropriately
    }
    const rec1 = ChangeExp.find(rc1 => rc1.level === parsedLevel);

    // Log the record if the level exists in the array
    if (rec1) {
      return rec1.expRatio;  // Return the expRatio
    } else {
      return 0;  // Return 0 if no record found
    }
  }

  const experienceFilename = 'global\\excel\\experience.txt';
  const experience = D2RMM.readTsv(experienceFilename);
  experience.rows.forEach((row) => {
    const lvl = row['Level'];
    const theChangeExp = findExpRatioByLevel(lvl);
    if (theChangeExp > 0) {
      console.log(`Changing level ${lvl} ExpRatio to ${theChangeExp}`);
      row['ExpRatio\r'] = theChangeExp.toString() + '\r';
    }
  });

  D2RMM.writeTsv(experienceFilename, experience);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation gamble.txt'
{
  const strArr = [
    { name: 'Wand', 'code\r': 'wnd\r' },
    { name: 'Yew Wand', 'code\r': 'ywn\r' },
    { name: 'Bone Wand', 'code\r': 'bwn\r' },
    { name: 'Grim Wand', 'code\r': 'gwn\r' },
    { name: 'Eagle Orb', 'code\r': 'ob1\r' },
    { name: 'Scepter', 'code\r': 'scp\r' },
    { name: 'Grand Scepter', 'code\r': 'gsc\r' },
    { name: 'War Scepter', 'code\r': 'wsp\r' },
    { name: 'Short Staff', 'code\r': 'sst\r' },
    { name: 'Long Staff', 'code\r': 'lst\r' },
    { name: 'Gnarled Staff', 'code\r': 'cst\r' },
    { name: 'Battle Staff', 'code\r': 'bst\r' },
    { name: 'War Staff', 'code\r': 'wst\r' },
    { name: 'Sacred Globe', 'code\r': 'ob2\r' },
    { name: 'Smoked Sphere', 'code\r': 'ob3\r' },
    { name: 'Clasped Orb', 'code\r': 'ob4\r' },
    { name: 'Dragon Stone', 'code\r': 'ob5\r' },
    { name: 'Stag Bow', 'code\r': 'am1\r' },
    { name: 'Reflex Bow', 'code\r': 'am2\r' },
    { name: 'Maiden Spear', 'code\r': 'am3\r' },
    { name: 'Maiden Pike', 'code\r': 'am4\r' },
    { name: 'Maiden Javelin', 'code\r': 'am5\r' },
    { name: 'Wolf Head', 'code\r': 'dr1\r' },
    { name: 'Hawk Helm', 'code\r': 'dr2\r' },
    { name: 'Antlers', 'code\r': 'dr3\r' },
    { name: 'Falcon Mask', 'code\r': 'dr4\r' },
    { name: 'Spirit Mask', 'code\r': 'dr5\r' },
    { name: 'Jawbone Cap', 'code\r': 'ba1\r' },
    { name: 'Fanged Helm', 'code\r': 'ba2\r' },
    { name: 'Horned Helm', 'code\r': 'ba3\r' },
    { name: 'Assault Helmet', 'code\r': 'ba4\r' },
    { name: 'Avenger Guard', 'code\r': 'ba5\r' },
    { name: 'Targe', 'code\r': 'pa1\r' },
    { name: 'Rondache', 'code\r': 'pa2\r' },
    { name: 'Heraldic Shield', 'code\r': 'pa3\r' },
    { name: 'Aerin Shield', 'code\r': 'pa4\r' },
    { name: 'Crown Shield', 'code\r': 'pa5\r' },
    { name: 'Preserved Head', 'code\r': 'ne1\r' },
    { name: 'Zombie Head', 'code\r': 'ne2\r' },
    { name: 'Unraveller Head', 'code\r': 'ne3\r' },
    { name: 'Gargoyle Head', 'code\r': 'ne4\r' },
    { name: 'Demon Head', 'code\r': 'ne5\r' }
  ];
  {
    const gambleFilename = 'global\\excel\\gamble.txt';
    const gamble = D2RMM.readTsv(gambleFilename);

    function addGambling(i) {
      gamble.rows.push({ ...i });
      D2RMM.writeTsv(gambleFilename, gamble);
    }

    strArr.forEach((s) => {
      addGambling(s);
    });
  }
}

// D2SE_Enjoy-SP_Mod_1.7 implementation shrines.txt'
{
  const shrinesFilename = 'global\\excel\\shrines.txt';
  const shrines = D2RMM.readTsv(shrinesFilename);
  shrines.rows.forEach((row) => {
    const theName = row['Name'];
    if (theName === 'Armor Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Combat Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Resist Fire Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Resist Cold Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Resist Lightning Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Resist Poison Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Skill Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Mana Recharge Shrine') {
      row['Duration in frames'] = 4800;
    }
    if (theName === 'Experience Shrine') {
      row['Duration in frames'] = 4800;
    }
  });
  D2RMM.writeTsv(shrinesFilename, shrines);
}

//TODO Skills
{
  //const skillsFilename = 'global\\excel\\skills.txt';
  //const skills = D2RMM.readTsv(skillsFilename);
  //skills.rows.forEach((row) => {
  //const sk = row['skill'];
  //if (sk === 'Volcano') {
  //console.log(row['localdelay']);
  //row['localdelay'] = 50;
  //console.log(row['localdelay']);
  //}
  //});
  //D2RMM.writeTsv(skillsFilename, skills);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation uniqueitems.txt'
{
  const NameAndRarity = [
    { name: "The Gnasher", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Deathspade", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Bladebone", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Mindrend", rarity: 1, nolimit: 1, level: 28, lvlReq: 21 },
    { name: "Rakescar", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Fechmars Axe", rarity: 1, nolimit: 1, level: 11, lvlReq: 8 },
    { name: "Goreshovel", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "The Chieftan", rarity: 1, nolimit: 1, level: 26, lvlReq: 19 },
    { name: "Brainhew", rarity: 1, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "The Humongous", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Iros Torch", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Maelstromwrath", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "Gravenspine", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Umes Lament", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Felloak", rarity: 1, nolimit: 1, level: 4, lvlReq: 3 },
    { name: "Knell Striker", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Rusthandle", rarity: 1, nolimit: 1, level: 23, lvlReq: 17 },
    { name: "Stormeye", rarity: 1, nolimit: 1, level: 31, lvlReq: 23 },
    { name: "Stoutnail", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Crushflange", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Bloodrise", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "The Generals Tan Do Li Ga", rarity: 1, nolimit: 1, level: 28, lvlReq: 21 },
    { name: "Ironstone", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Bonesob", rarity: 1, nolimit: 1, level: 32, lvlReq: 24 },
    { name: "Steeldriver", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Rixots Keen", rarity: 1, nolimit: 1, level: 3, lvlReq: 2 },
    { name: "Blood Crescent", rarity: 1, nolimit: 1, level: 10, lvlReq: 7 },
    { name: "Krintizs Skewer", rarity: 1, nolimit: 1, level: 14, lvlReq: 10 },
    { name: "Gleamscythe", rarity: 1, nolimit: 1, level: 18, lvlReq: 13 },
    { name: "Azurewrath", rarity: 1, nolimit: 1, level: 18, lvlReq: 13 },
    { name: "Griswolds Edge", rarity: 1, nolimit: 1, level: 8, lvlReq: 14 },
    { name: "Hellplague", rarity: 1, nolimit: 1, level: 30, lvlReq: 22 },
    { name: "Culwens Point", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Shadowfang", rarity: 1, nolimit: 1, level: 16, lvlReq: 12 },
    { name: "Soulflay", rarity: 1, nolimit: 1, level: 20, lvlReq: 19 },
    { name: "Kinemils Awl", rarity: 1, nolimit: 1, level: 31, lvlReq: 23 },
    { name: "Blacktongue", rarity: 1, nolimit: 1, level: 35, lvlReq: 26 },
    { name: "Ripsaw", rarity: 1, nolimit: 1, level: 35, lvlReq: 26 },
    { name: "The Patriarch", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Gull", rarity: 1, nolimit: 1, level: 6, lvlReq: 4 },
    { name: "The Diggler", rarity: 1, nolimit: 1, level: 15, lvlReq: 11 },
    { name: "The Jade Tan Do", rarity: 1, nolimit: 1, level: 26, lvlReq: 19 },
    { name: "Irices Shard", rarity: 1, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "The Dragon Chang", rarity: 1, nolimit: 1, level: 11, lvlReq: 8 },
    { name: "Razortine", rarity: 1, nolimit: 1, level: 16, lvlReq: 12 },
    { name: "Bloodthief", rarity: 1, nolimit: 1, level: 23, lvlReq: 17 },
    { name: "Lance of Yaggai", rarity: 1, nolimit: 1, level: 30, lvlReq: 22 },
    { name: "The Tannr Gorerod", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Dimoaks Hew", rarity: 1, nolimit: 1, level: 11, lvlReq: 8 },
    { name: "Steelgoad", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "Soul Harvest", rarity: 1, nolimit: 1, level: 26, lvlReq: 19 },
    { name: "The Battlebranch", rarity: 1, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "Woestave", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "The Grim Reaper", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Bane Ash", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Serpent Lord", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Lazarus Spire", rarity: 1, nolimit: 1, level: 24, lvlReq: 18 },
    { name: "The Salamander", rarity: 1, nolimit: 1, level: 28, lvlReq: 21 },
    { name: "The Iron Jang Bong", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Pluckeye", rarity: 1, nolimit: 1, level: 10, lvlReq: 7 },
    { name: "Witherstring", rarity: 1, nolimit: 1, level: 18, lvlReq: 13 },
    { name: "Rimeraven", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Piercerib", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Pullspite", rarity: 1, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "Wizendraw", rarity: 1, nolimit: 1, level: 35, lvlReq: 26 },
    { name: "Hellclap", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Blastbark", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Leadcrow", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Ichorsting", rarity: 1, nolimit: 1, level: 24, lvlReq: 18 },
    { name: "Hellcast", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Doomspittle", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "War Bonnet", rarity: 1, nolimit: 1, level: 4, lvlReq: 3 },
    { name: "Tarnhelm", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Coif of Glory", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "Duskdeep", rarity: 1, nolimit: 1, level: 23, lvlReq: 17 },
    { name: "Wormskull", rarity: 1, nolimit: 1, level: 28, lvlReq: 21 },
    { name: "Howltusk", rarity: 1, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "Undead Crown", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "The Face of Horror", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Greyform", rarity: 1, nolimit: 1, level: 10, lvlReq: 7 },
    { name: "Blinkbats Form", rarity: 1, nolimit: 1, level: 16, lvlReq: 12 },
    { name: "The Centurion", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "Twitchthroe", rarity: 1, nolimit: 1, level: 22, lvlReq: 16 },
    { name: "Darkglow", rarity: 1, nolimit: 1, level: 19, lvlReq: 14 },
    { name: "Hawkmail", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Sparking Mail", rarity: 1, nolimit: 1, level: 23, lvlReq: 17 },
    { name: "Venomsward", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Iceblink", rarity: 1, nolimit: 1, level: 30, lvlReq: 22 },
    { name: "Boneflesh", rarity: 1, nolimit: 1, level: 35, lvlReq: 26 },
    { name: "Rockfleece", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Rattlecage", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Goldskin", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Victors Silk", rarity: 1, nolimit: 1, level: 38, lvlReq: 28 },
    { name: "Heavenly Garb", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Pelta Lunata", rarity: 1, nolimit: 1, level: 3, lvlReq: 2 },
    { name: "Umbral Disk", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Stormguild", rarity: 1, nolimit: 1, level: 18, lvlReq: 13 },
    { name: "Wall of the Eyeless", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Swordback Hold", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Steelclash", rarity: 1, nolimit: 1, level: 23, lvlReq: 17 },
    { name: "Bverrit Keep", rarity: 1, nolimit: 1, level: 26, lvlReq: 19 },
    { name: "The Ward", rarity: 1, nolimit: 1, level: 35, lvlReq: 26 },
    { name: "The Hand of Broc", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Bloodfist", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Chance Guards", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Magefist", rarity: 1, nolimit: 1, level: 31, lvlReq: 23 },
    { name: "Frostburn", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Hotspur", rarity: 1, nolimit: 1, level: 7, lvlReq: 5 },
    { name: "Gorefoot", rarity: 1, nolimit: 1, level: 12, lvlReq: 9 },
    { name: "Treads of Cthon", rarity: 1, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "Goblin Toe", rarity: 1, nolimit: 1, level: 30, lvlReq: 22 },
    { name: "Tearhaunch", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Lenyms Cord", rarity: 1, nolimit: 1, level: 10, lvlReq: 7 },
    { name: "Snakecord", rarity: 1, nolimit: 1, level: 16, lvlReq: 12 },
    { name: "Nightsmoke", rarity: 1, nolimit: 1, level: 27, lvlReq: 20 },
    { name: "Goldwrap", rarity: 1, nolimit: 1, level: 36, lvlReq: 27 },
    { name: "Bladebuckle", rarity: 1, nolimit: 1, level: 39, lvlReq: 29 },
    { name: "Nokozan Relic", rarity: 20, nolimit: 1, level: 14, lvlReq: 10 },
    { name: "The Eye of Etlich", rarity: 5, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "The Mahim-Oak Curio", rarity: 10, nolimit: 1, level: 34, lvlReq: 25 },
    { name: "Nagelring", rarity: 15, nolimit: 1, level: 10, lvlReq: 7 },
    { name: "Manald Heal", rarity: 15, nolimit: 1, level: 20, lvlReq: 15 },
    { name: "The Stone of Jordan", rarity: 5, nolimit: 1, level: 41, lvlReq: 40 },
    { name: "Amulet of the Viper", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "Staff of Kings", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "Horadric Staff", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "Hell Forge Hammer", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "KhalimFlail", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "SuperKhalimFlail", rarity: 1, nolimit: 1, level: 0, lvlReq: 0 },
    { name: "Coldkill", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Butcher's Pupil", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Islestrike", rarity: 1, nolimit: 1, level: 51, lvlReq: 43 },
    { name: "Pompe's Wrath", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Guardian Naga", rarity: 1, nolimit: 1, level: 56, lvlReq: 48 },
    { name: "Warlord's Trust", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Spellsteel", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Stormrider", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Boneslayer Blade", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "The Minataur", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Suicide Branch", rarity: 1, nolimit: 1, level: 41, lvlReq: 33 },
    { name: "Carin Shard", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Arm of King Leoric", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Blackhand Key", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Dark Clan Crusher", rarity: 1, nolimit: 1, level: 42, lvlReq: 34 },
    { name: "Zakarum's Hand", rarity: 1, nolimit: 1, level: 45, lvlReq: 37 },
    { name: "The Fetid Sprinkler", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Hand of Blessed Light", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Fleshrender", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Sureshrill Frost", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Moonfall", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Baezil's Vortex", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Earthshaker", rarity: 1, nolimit: 1, level: 51, lvlReq: 43 },
    { name: "Bloodtree Stump", rarity: 1, nolimit: 1, level: 56, lvlReq: 48 },
    { name: "The Gavel of Pain", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Bloodletter", rarity: 1, nolimit: 1, level: 38, lvlReq: 30 },
    { name: "Coldsteel Eye", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Hexfire", rarity: 1, nolimit: 1, level: 41, lvlReq: 33 },
    { name: "Blade of Ali Baba", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Ginther's Rift", rarity: 1, nolimit: 1, level: 45, lvlReq: 37 },
    { name: "Headstriker", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Plague Bearer", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "The Atlantian", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Crainte Vomir", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Bing Sz Wang", rarity: 1, nolimit: 1, level: 51, lvlReq: 43 },
    { name: "The Vile Husk", rarity: 1, nolimit: 1, level: 52, lvlReq: 44 },
    { name: "Cloudcrack", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Todesfaelle Flamme", rarity: 1, nolimit: 1, level: 54, lvlReq: 46 },
    { name: "Swordguard", rarity: 1, nolimit: 1, level: 55, lvlReq: 48 },
    { name: "Spineripper", rarity: 1, nolimit: 1, level: 40, lvlReq: 32 },
    { name: "Heart Carver", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Blackbog's Sharp", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Stormspike", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "The Impaler", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Kelpie Snare", rarity: 1, nolimit: 1, level: 41, lvlReq: 33 },
    { name: "Soulfeast Tine", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Hone Sundan", rarity: 1, nolimit: 1, level: 45, lvlReq: 37 },
    { name: "Spire of Honor", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "The Meat Scraper", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Blackleach Blade", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Athena's Wrath", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Pierre Tombale Couant", rarity: 1, nolimit: 1, level: 51, lvlReq: 43 },
    { name: "Husoldal Evo", rarity: 1, nolimit: 1, level: 52, lvlReq: 44 },
    { name: "Grim's Burning Dead", rarity: 1, nolimit: 1, level: 52, lvlReq: 45 },
    { name: "Razorswitch", rarity: 1, nolimit: 1, level: 36, lvlReq: 28 },
    { name: "Ribcracker", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Chromatic Ire", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Warpspear", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Skullcollector", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Skystrike", rarity: 1, nolimit: 1, level: 36, lvlReq: 28 },
    { name: "Riphook", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Kuko Shakaku", rarity: 1, nolimit: 1, level: 41, lvlReq: 33 },
    { name: "Endlesshail", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Whichwild String", rarity: 1, nolimit: 1, level: 47, lvlReq: 39 },
    { name: "Cliffkiller", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Magewrath", rarity: 1, nolimit: 1, level: 51, lvlReq: 43 },
    { name: "Godstrike Arch", rarity: 1, nolimit: 1, level: 54, lvlReq: 46 },
    { name: "Langer Briser", rarity: 1, nolimit: 1, level: 40, lvlReq: 32 },
    { name: "Pus Spiter", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Buriza-Do Kyanon", rarity: 1, nolimit: 1, level: 59, lvlReq: 41 },
    { name: "Demon Machine", rarity: 1, nolimit: 1, level: 57, lvlReq: 49 },
    { name: "Peasent Crown", rarity: 1, nolimit: 1, level: 36, lvlReq: 28 },
    { name: "Rockstopper", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Stealskull", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Darksight Helm", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Valkiry Wing", rarity: 1, nolimit: 1, level: 52, lvlReq: 44 },
    { name: "Crown of Thieves", rarity: 1, nolimit: 1, level: 57, lvlReq: 49 },
    { name: "Blackhorn's Face", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Vampiregaze", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "The Spirit Shroud", rarity: 1, nolimit: 1, level: 36, lvlReq: 28 },
    { name: "Skin of the Vipermagi", rarity: 1, nolimit: 1, level: 37, lvlReq: 29 },
    { name: "Skin of the Flayerd One", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Ironpelt", rarity: 1, nolimit: 1, level: 41, lvlReq: 33 },
    { name: "Spiritforge", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Crow Caw", rarity: 1, nolimit: 1, level: 45, lvlReq: 37 },
    { name: "Shaftstop", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Duriel's Shell", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Skullder's Ire", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Guardian Angel", rarity: 1, nolimit: 1, level: 53, lvlReq: 45 },
    { name: "Toothrow", rarity: 1, nolimit: 1, level: 56, lvlReq: 48 },
    { name: "Atma's Wail", rarity: 1, nolimit: 1, level: 59, lvlReq: 51 },
    { name: "Black Hades", rarity: 1, nolimit: 1, level: 61, lvlReq: 53 },
    { name: "Corpsemourn", rarity: 1, nolimit: 1, level: 63, lvlReq: 55 },
    { name: "Que-Hegan's Wisdon", rarity: 1, nolimit: 1, level: 59, lvlReq: 51 },
    { name: "Visceratuant", rarity: 1, nolimit: 1, level: 36, lvlReq: 28 },
    { name: "Mosers Blessed Circle", rarity: 1, nolimit: 1, level: 39, lvlReq: 31 },
    { name: "Stormchaser", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Tiamat's Rebuke", rarity: 1, nolimit: 1, level: 46, lvlReq: 38 },
    { name: "Kerke's Sanctuary", rarity: 1, nolimit: 1, level: 52, lvlReq: 44 },
    { name: "Radimant's Sphere", rarity: 1, nolimit: 1, level: 58, lvlReq: 50 },
    { name: "Lidless Wall", rarity: 1, nolimit: 1, level: 49, lvlReq: 41 },
    { name: "Lance Guard", rarity: 1, nolimit: 1, level: 43, lvlReq: 35 },
    { name: "Venom Grip", rarity: 1, nolimit: 1, level: 37, lvlReq: 29 },
    { name: "Gravepalm", rarity: 1, nolimit: 1, level: 39, lvlReq: 32 },
    { name: "Ghoulhide", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Lavagout", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Hellmouth", rarity: 1, nolimit: 1, level: 55, lvlReq: 47 },
    { name: "Infernostride", rarity: 1, nolimit: 1, level: 37, lvlReq: 29 },
    { name: "Waterwalk", rarity: 1, nolimit: 1, level: 40, lvlReq: 32 },
    { name: "Silkweave", rarity: 1, nolimit: 1, level: 44, lvlReq: 36 },
    { name: "Wartraveler", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Gorerider", rarity: 1, nolimit: 1, level: 55, lvlReq: 47 },
    { name: "String of Ears", rarity: 1, nolimit: 1, level: 37, lvlReq: 29 },
    { name: "Razortail", rarity: 1, nolimit: 1, level: 39, lvlReq: 32 },
    { name: "Gloomstrap", rarity: 1, nolimit: 1, level: 45, lvlReq: 36 },
    { name: "Snowclash", rarity: 1, nolimit: 1, level: 49, lvlReq: 42 },
    { name: "Thudergod's Vigor", rarity: 1, nolimit: 1, level: 55, lvlReq: 47 },
    { name: "Harlequin Crest", rarity: 1, nolimit: 1, level: 62, lvlReq: 62 },
    { name: "Veil of Steel", rarity: 1, nolimit: 1, level: 64, lvlReq: 64 },
    { name: "The Gladiator's Bane", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Arkaine's Valor", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Blackoak Shield", rarity: 1, nolimit: 1, level: 66, lvlReq: 61 },
    { name: "Stormshield", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Hellslayer", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Messerschmidt's Reaver", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Baranar's Star", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Schaefer's Hammer", rarity: 1, nolimit: 1, level: 69, lvlReq: 67 },
    { name: "The Cranium Basher", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Lightsabre", rarity: 1, nolimit: 1, level: 66, lvlReq: 58 },
    { name: "Doombringer", rarity: 1, nolimit: 1, level: 64, lvlReq: 64 },
    { name: "The Grandfather", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Wizardspike", rarity: 1, nolimit: 1, level: 61, lvlReq: 61 },
    { name: "Constricting Ring", rarity: 1, nolimit: 1, level: 110, lvlReq: 90 },
    { name: "Stormspire", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Eaglehorn", rarity: 1, nolimit: 1, level: 62, lvlReq: 62 },
    { name: "Windforce", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Bul Katho's Wedding Band", rarity: 7, nolimit: 1, level: 61, lvlReq: 58 },
    { name: "The Cat's Eye", rarity: 12, nolimit: 1, level: 50, lvlReq: 50 },
    { name: "The Rising Sun", rarity: 12, nolimit: 1, level: 60, lvlReq: 60 },
    { name: "Crescent Moon", rarity: 14, nolimit: 1, level: 50, lvlReq: 50 },
    { name: "Mara's Kaleidoscope", rarity: 8, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Atma's Scarab", rarity: 12, nolimit: 1, level: 55, lvlReq: 55 },
    { name: "Dwarf Star", rarity: 5, nolimit: 1, level: 8, lvlReq: 8 },
    { name: "Raven Frost", rarity: 12, nolimit: 1, level: 46, lvlReq: 40 },
    { name: "Highlord's Wrath", rarity: 9, nolimit: 1, level: 60, lvlReq: 60 },
    { name: "Saracen's Chance", rarity: 14, nolimit: 1, level: 45, lvlReq: 45 },
    { name: "Arreat's Face", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Homunculus", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Titan's Revenge", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Lycander's Aim", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Lycander's Flank", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "The Oculus", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Herald of Zakarum", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Cutthroat1", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "Jalal's Mane", rarity: 1, nolimit: 1, level: 50, lvlReq: 42 },
    { name: "The Scalper", rarity: 1, nolimit: 1, level: 65, lvlReq: 57 },
    { name: "Bloodmoon", rarity: 1, nolimit: 1, level: 62, lvlReq: 61 },
    { name: "Djinnslayer", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Deathbit", rarity: 1, nolimit: 1, level: 52, lvlReq: 44 },
    { name: "Warshrike", rarity: 1, nolimit: 1, level: 69, lvlReq: 69 },
    { name: "Gutsiphon", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Razoredge", rarity: 1, nolimit: 1, level: 66, lvlReq: 67 },
    { name: "Gore Ripper", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Demonlimb", rarity: 1, nolimit: 1, level: 66, lvlReq: 63 },
    { name: "Steelshade", rarity: 1, nolimit: 1, level: 65, lvlReq: 62 },
    { name: "Tomb Reaver", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Deaths's Web", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Nature's Peace", rarity: 9, nolimit: 1, level: 58, lvlReq: 58 },
    { name: "Azurewrath", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Seraph's Hymn", rarity: 12, nolimit: 1, level: 66, lvlReq: 65 },
    { name: "Zakarum's Salvation", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Fleshripper", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Odium", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Horizon's Tornado", rarity: 2, nolimit: 1, level: 66, lvlReq: 64 },
    { name: "Stone Crusher", rarity: 1, nolimit: 1, level: 66, lvlReq: 68 },
    { name: "Jadetalon", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Shadowdancer", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Cerebus", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Tyrael's Might", rarity: 1, nolimit: 1, level: 72, lvlReq: 72 },
    { name: "Souldrain", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Runemaster", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Deathcleaver", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Executioner's Justice", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Stoneraven", rarity: 1, nolimit: 1, level: 66, lvlReq: 64 },
    { name: "Leviathan", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Larzuk's Champion", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Wisp", rarity: 9, nolimit: 1, level: 55, lvlReq: 45 },
    { name: "Gargoyle's Bite", rarity: 1, nolimit: 1, level: 58, lvlReq: 58 },
    { name: "Lacerator", rarity: 1, nolimit: 1, level: 66, lvlReq: 68 },
    { name: "Mang Song's Lesson", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Viperfork", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Ethereal Edge", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Demonhorn's Edge", rarity: 1, nolimit: 1, level: 65, lvlReq: 55 },
    { name: "The Reaper's Toll", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Spiritkeeper", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Hellrack", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Alma Negra", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Darkforge Spawn", rarity: 1, nolimit: 1, level: 64, lvlReq: 64 },
    { name: "Widowmaker", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Bloodraven's Charge", rarity: 1, nolimit: 1, level: 58, lvlReq: 58 },
    { name: "Ghostflame", rarity: 1, nolimit: 1, level: 63, lvlReq: 62 },
    { name: "Shadowkiller", rarity: 1, nolimit: 1, level: 69, lvlReq: 69 },
    { name: "Gimmershred", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Griffon's Eye", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Windhammer", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Thunderstroke", rarity: 1, nolimit: 1, level: 66, lvlReq: 69 },
    { name: "Giantmaimer", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Demon's Arch", rarity: 1, nolimit: 1, level: 66, lvlReq: 68 },
    { name: "Boneflame", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Steelpillar", rarity: 1, nolimit: 1, level: 67, lvlReq: 69 },
    { name: "Nightwing's Veil", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Crown of Ages", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Andariel's Visage", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Darkfear", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Dragonscale", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Steel Carapice", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Medusa's Gaze", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Ravenlore", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Boneshade", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Nethercrow", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Flamebellow", rarity: 1, nolimit: 1, level: 79, lvlReq: 71 },
    { name: "Fathom", rarity: 1, nolimit: 1, level: 71, lvlReq: 73 },
    { name: "Wolfhowl", rarity: 1, nolimit: 1, level: 69, lvlReq: 79 },
    { name: "Spirit Ward", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Kira's Guardian", rarity: 1, nolimit: 1, level: 69, lvlReq: 77 },
    { name: "Ormus' Robes", rarity: 1, nolimit: 1, level: 69, lvlReq: 75 },
    { name: "Gheed's Fortune", rarity: 1, nolimit: 1, level: 70, lvlReq: 62 },
    { name: "Stormlash", rarity: 1, nolimit: 1, level: 69, lvlReq: 69 },
    { name: "Halaberd's Reign", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Warriv's Warder", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Spike Thorn", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Dracul's Grasp", rarity: 1, nolimit: 1, level: 66, lvlReq: 67 },
    { name: "Frostwind", rarity: 1, nolimit: 1, level: 64, lvlReq: 64 },
    { name: "Templar's Might", rarity: 3, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Eschuta's temper", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Firelizard's Talons", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Sandstorm Trek", rarity: 1, nolimit: 1, level: 65, lvlReq: 64 },
    { name: "Marrowwalk", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "Heaven's Light", rarity: 1, nolimit: 1, level: 63, lvlReq: 61 },
    { name: "Arachnid Mesh", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Nosferatu's Coil", rarity: 1, nolimit: 1, level: 67, lvlReq: 51 },
    { name: "Metalgrid", rarity: 11, nolimit: 1, level: 56, lvlReq: 56 },
    { name: "Verdugo's Hearty Cord", rarity: 1, nolimit: 1, level: 65, lvlReq: 63 },
    { name: "Sigurd's Staunch", rarity: 1, nolimit: 1, level: '', lvlReq: '' },
    { name: "Carrion Wind", rarity: 8, nolimit: 1, level: 56, lvlReq: 50 },
    { name: "Giantskull", rarity: 1, nolimit: 1, level: 65, lvlReq: 65 },
    { name: "Ironward", rarity: 1, nolimit: 1, level: 61, lvlReq: 60 },
    { name: "Annihilus", rarity: 1, nolimit: 1, level: 110, lvlReq: 70 },
    { name: "Arioc's Needle", rarity: 1, nolimit: 1, level: 69, lvlReq: 69 },
    { name: "Cranebeak", rarity: 1, nolimit: 1, level: 66, lvlReq: 63 },
    { name: "Nord's Tenderizer", rarity: 1, nolimit: 1, level: 66, lvlReq: 68 },
    { name: "Earthshifter", rarity: 1, nolimit: 1, level: 66, lvlReq: 69 },
    { name: "Wraithflight", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Bonehew", rarity: 1, nolimit: 1, level: 65, lvlReq: 64 },
    { name: "Ondal's Wisdom", rarity: 1, nolimit: 1, level: 66, lvlReq: 66 },
    { name: "The Reedeemer", rarity: 1, nolimit: 1, level: 67, lvlReq: 67 },
    { name: "Headhunter's Glory", rarity: 1, nolimit: 1, level: 68, lvlReq: 68 },
    { name: "Steelrend", rarity: 1, nolimit: 1, level: 70, lvlReq: 70 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Rainbow Facet", rarity: 1, nolimit: 1, level: 50, lvlReq: 49 },
    { name: "Hellfire Torch", rarity: 1, nolimit: 1, level: 110, lvlReq: 75 },
  ];

  function findRarityByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.rarity : null; // Return null if not found
  }
  function findLevelByLevel(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.level : null; // Return null if not found
  }
  function findLevelReqByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.lvlReq : null; // Return null if not found
  }
  function findNoLimitByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.nolimit : null; // Return null if not found
  }

  const uniqueitemsFilename = 'global\\excel\\uniqueitems.txt';
  const uniqueitems = D2RMM.readTsv(uniqueitemsFilename);
  uniqueitems.rows.forEach((row) => {
    const theName = row['index'];
    const rarity = findRarityByName(theName);
    const nolimit = findNoLimitByName(theName);
    const level = findLevelByLevel(theName);
    const lvlReq = findLevelReqByName(theName);
    if (lvlReq !== null) {
      row['lvl req'] = lvlReq;
    }
    if (rarity !== null) {
      row.rarity = rarity;
    }
    if (nolimit !== null) {
      row.nolimit = nolimit;
    }
    if (level !== null) {
      row.lvl = level;
    }
  });
  D2RMM.writeTsv(uniqueitemsFilename, uniqueitems);
}

// D2SE_Enjoy-SP_Mod_1.7 implementation weapons.txt'
{
  const NameAndRarity = [
    { name: "Hand Axe", rarity: 1, level: 3 },
    { name: "Axe", rarity: 1, level: 7 },
    { name: "Double Axe", rarity: 1, level: 13 },
    { name: "Military Pick", rarity: 1, level: 19 },
    { name: "War Axe", rarity: 1, level: 25 },
    { name: "Large Axe", rarity: 1, level: 6 },
    { name: "Broad Axe ", rarity: 1, level: 12 },
    { name: "Battle Axe", rarity: 1, level: 17 },
    { name: "Great Axe", rarity: 2, level: 23 },
    { name: "Giant Axe", rarity: 2, level: 27 },
    { name: "Wand", rarity: 1, level: 2 },
    { name: "Yew Wand", rarity: 1, level: 12 },
    { name: "Bone Wand", rarity: 1, level: 18 },
    { name: "Grim Wand", rarity: 1, level: 26 },
    { name: "Club", rarity: 1, level: 1 },
    { name: "Scepter", rarity: 1, level: 3 },
    { name: "Grand Scepter", rarity: 2, level: 15 },
    { name: "War Scepter", rarity: 3, level: 21 },
    { name: "Spiked Club", rarity: 1, level: 4 },
    { name: "Mace", rarity: 1, level: 8 },
    { name: "Morning Star", rarity: 1, level: 13 },
    { name: "Flail", rarity: 2, level: 19 },
    { name: "War Hammer", rarity: 2, level: 25 },
    { name: "Maul", rarity: 1, level: 21 },
    { name: "Great Maul", rarity: 1, level: 32 },
    { name: "Short Sword", rarity: 1, level: 1 },
    { name: "Scimitar", rarity: 1, level: 5 },
    { name: "Saber", rarity: 1, level: 8 },
    { name: "Falchion", rarity: 2, level: 11 },
    { name: "Crystal Sword", rarity: 1, level: 11 },
    { name: "Broad Sword", rarity: 2, level: 15 },
    { name: "Long Sword ", rarity: 1, level: 20 },
    { name: "War Sword", rarity: 1, level: 27 },
    { name: "Two-Handed Sword", rarity: 2, level: 10 },
    { name: "Claymore", rarity: 2, level: 17 },
    { name: "Giant Sword", rarity: 2, level: 21 },
    { name: "Bastard Sword", rarity: 3, level: 24 },
    { name: "Flamberge", rarity: 2, level: 27 },
    { name: "Great Sword", rarity: 2, level: 33 },
    { name: "Dagger", rarity: 1, level: 3 },
    { name: "Dirk", rarity: 1, level: 9 },
    { name: "Kriss", rarity: 1, level: 17 },
    { name: "Blade", rarity: 1, level: 23 },
    { name: "Throwing Knife", rarity: 1, level: 2 },
    { name: "Throwing Axe", rarity: 1, level: 7 },
    { name: "Balanced Knife", rarity: 1, level: 13 },
    { name: "Balanced Axe", rarity: 1, level: 16 },
    { name: "Javelin", rarity: 1, level: 1 },
    { name: "Pilum", rarity: 1, level: 10 },
    { name: "Short Spear", rarity: 1, level: 15 },
    { name: "Glaive", rarity: 1, level: 23 },
    { name: "Throwing Spear", rarity: 1, level: 29 },
    { name: "Spear", rarity: 2, level: 5 },
    { name: "Trident", rarity: 2, level: 9 },
    { name: "Brandistock", rarity: 2, level: 16 },
    { name: "Spetum", rarity: 2, level: 20 },
    { name: "Pike", rarity: 2, level: 24 },
    { name: "Bardiche", rarity: 2, level: 5 },
    { name: "Voulge", rarity: 2, level: 11 },
    { name: "Scythe", rarity: 2, level: 15 },
    { name: "Poleaxe", rarity: 2, level: 21 },
    { name: "Halberd", rarity: 2, level: 29 },
    { name: "War Scythe", rarity: 2, level: 34 },
    { name: "Short Staff", rarity: 2, level: 1 },
    { name: "Long Staff", rarity: 2, level: 8 },
    { name: "Gnarled Staff", rarity: 2, level: 12 },
    { name: "Battle Staff", rarity: 2, level: 17 },
    { name: "War Staff", rarity: 2, level: 24 },
    { name: "Short Bow", rarity: 2, level: 1 },
    { name: "Hunter's Bow", rarity: 2, level: 5 },
    { name: "Long Bow", rarity: 2, level: 8 },
    { name: "Composite Bow", rarity: 2, level: 12 },
    { name: "Short Battle Bow", rarity: 2, level: 18 },
    { name: "Long Battle Bow", rarity: 2, level: 23 },
    { name: "Short War Bow", rarity: 2, level: 27 },
    { name: "Long War Bow", rarity: 2, level: 31 },
    { name: "Light Crossbow", rarity: 2, level: 6 },
    { name: "Crossbow", rarity: 2, level: 15 },
    { name: "Heavy Crossbow", rarity: 2, level: 24 },
    { name: "Repeating Crossbow", rarity: 2, level: 33 },
    { name: "Hatchet", rarity: 1, level: 31 },
    { name: "Cleaver", rarity: 1, level: 34 },
    { name: "Twin Axe", rarity: 1, level: 39 },
    { name: "Crowbill", rarity: 1, level: 43 },
    { name: "Naga", rarity: 1, level: 48 },
    { name: "Military Axe", rarity: 1, level: 34 },
    { name: "Bearded Axe ", rarity: 1, level: 38 },
    { name: "Tabar", rarity: 1, level: 42 },
    { name: "Gothic Axe", rarity: 1, level: 46 },
    { name: "Ancient Axe", rarity: 2, level: 51 },
    { name: "Burnt Wand", rarity: 1, level: 31 },
    { name: "Petrified Wand", rarity: 1, level: 38 },
    { name: "Tomb Wand", rarity: 1, level: 43 },
    { name: "Grave Wand", rarity: 1, level: 49 },
    { name: "Cudgel", rarity: 1, level: 30 },
    { name: "Rune Scepter", rarity: 2, level: 31 },
    { name: "Holy Water Sprinkler", rarity: 2, level: 40 },
    { name: "Divine Scepter", rarity: 3, level: 45 },
    { name: "Barbed Club", rarity: 1, level: 32 },
    { name: "Flanged Mace", rarity: 2, level: 35 },
    { name: "Jagged Star", rarity: 2, level: 39 },
    { name: "Knout", rarity: 1, level: 43 },
    { name: "Battle Hammer", rarity: 2, level: 48 },
    { name: "War Club", rarity: 1, level: 45 },
    { name: "Martel de Fer", rarity: 1, level: 53 },
    { name: "Gladius", rarity: 1, level: 30 },
    { name: "Cutlass", rarity: 1, level: 43 },
    { name: "Shamshir", rarity: 1, level: 35 },
    { name: "Tulwar", rarity: 2, level: 37 },
    { name: "Dimensional Blade", rarity: 2, level: 37 },
    { name: "Battle Sword", rarity: 1, level: 40 },
    { name: "Rune Sword", rarity: 2, level: 44 },
    { name: "Ancient Sword", rarity: 2, level: 49 },
    { name: "Espandon", rarity: 2, level: 37 },
    { name: "Dacian Falx", rarity: 1, level: 42 },
    { name: "Tusk Sword", rarity: 1, level: 45 },
    { name: "Gothic Sword", rarity: 1, level: 48 },
    { name: "Zweihander", rarity: 1, level: 49 },
    { name: "Executioner Sword", rarity: 1, level: 54 },
    { name: "Poignard", rarity: 1, level: 31 },
    { name: "Rondel", rarity: 1, level: 36 },
    { name: "Cinquedeas", rarity: 1, level: 42 },
    { name: "Stilleto", rarity: 1, level: 46 },
    { name: "Battle Dart", rarity: 1, level: 31 },
    { name: "Francisca", rarity: 1, level: 34 },
    { name: "War Dart", rarity: 1, level: 39 },
    { name: "Hurlbat", rarity: 1, level: 41 },
    { name: "War Javelin", rarity: 2, level: 30 },
    { name: "Great Pilum", rarity: 2, level: 37 },
    { name: "Simbilan", rarity: 2, level: 40 },
    { name: "Spiculum", rarity: 2, level: 46 },
    { name: "Harpoon", rarity: 2, level: 51 },
    { name: "War Spear", rarity: 2, level: 33 },
    { name: "Fuscina", rarity: 2, level: 36 },
    { name: "War Fork", rarity: 2, level: 41 },
    { name: "Yari", rarity: 2, level: 44 },
    { name: "Lance", rarity: 2, level: 47 },
    { name: "Lochaber Axe", rarity: 2, level: 33 },
    { name: "Bill", rarity: 2, level: 37 },
    { name: "Battle Scythe", rarity: 2, level: 40 },
    { name: "Partizan", rarity: 2, level: 35 },
    { name: "Bec-de-Corbin", rarity: 2, level: 51 },
    { name: "Grim Scythe", rarity: 2, level: 55 },
    { name: "Jo Staff", rarity: 2, level: 30 },
    { name: "Quarterstaff", rarity: 2, level: 35 },
    { name: "Cedar Staff", rarity: 2, level: 38 },
    { name: "Gothic Staff", rarity: 2, level: 42 },
    { name: "Rune Staff", rarity: 2, level: 47 },
    { name: "Edge Bow", rarity: 2, level: 30 },
    { name: "Razor Bow", rarity: 2, level: 33 },
    { name: "Cedar Bow", rarity: 2, level: 35 },
    { name: "Double Bow", rarity: 2, level: 39 },
    { name: "Short Siege Bow", rarity: 2, level: 43 },
    { name: "Long Siege Bow", rarity: 2, level: 46 },
    { name: "Rune Bow", rarity: 2, level: 49 },
    { name: "Gothic Bow", rarity: 2, level: 52 },
    { name: "Arbalest", rarity: 2, level: 34 },
    { name: "Siege Crossbow", rarity: 2, level: 40 },
    { name: "Ballista", rarity: 2, level: 47 },
    { name: "Chu-Ko-Nu", rarity: 2, level: 54 },
    { name: "Katar", rarity: 1, level: 1 },
    { name: "Wrist Blade", rarity: 1, level: 9 },
    { name: "Hatchet Hands", rarity: 1, level: 12 },
    { name: "Cestus", rarity: 1, level: 15 },
    { name: "Claws", rarity: 1, level: 18 },
    { name: "Blade Talons", rarity: 1, level: 21 },
    { name: "Scissors Katar", rarity: 1, level: 24 },
    { name: "Quhab", rarity: 1, level: 28 },
    { name: "Wrist Spike", rarity: 1, level: 32 },
    { name: "Fascia", rarity: 2, level: 36 },
    { name: "Hand Scythe", rarity: 2, level: 41 },
    { name: "Greater Claws", rarity: 2, level: 45 },
    { name: "Greater Talons", rarity: 3, level: 50 },
    { name: "Scissors Quhab", rarity: 2, level: 54 },
    { name: "Suwayyah", rarity: 2, level: 59 },
    { name: "Wrist Sword", rarity: 2, level: 62 },
    { name: "War Fist", rarity: 2, level: 68 },
    { name: "Battle Cestus", rarity: 3, level: 66 },
    { name: "Feral Claws", rarity: 2, level: 66 },
    { name: "Runic Talons", rarity: 3, level: 66 },
    { name: "Scissors Suwayyah", rarity: 2, level: 66 },
    { name: "Tomahawk", rarity: 2, level: 54 },
    { name: "Small Crescent", rarity: 2, level: 61 },
    { name: "Ettin Axe", rarity: 2, level: 66 },
    { name: "War Spike", rarity: 2, level: 66 },
    { name: "Berserker Axe", rarity: 3, level: 66 },
    { name: "Feral Axe", rarity: 2, level: 57 },
    { name: "Silver-edged Axe", rarity: 2, level: 65 },
    { name: "Decapitator", rarity: 3, level: 56 },
    { name: "Champion Axe", rarity: 3, level: 66 },
    { name: "Glorious Axe", rarity: 3, level: 66 },
    { name: "Polished Wand", rarity: 1, level: 55 },
    { name: "Ghost Wand", rarity: 1, level: 56 },
    { name: "Lich Wand", rarity: 1, level: 56 },
    { name: "Unearthed Wand", rarity: 1, level: 66 },
    { name: "Truncheon", rarity: 1, level: 52 },
    { name: "Mighty Scepter", rarity: 2, level: 62 },
    { name: "Seraph Rod", rarity: 2, level: 66 },
    { name: "Caduceus", rarity: 2, level: 66 },
    { name: "Tyrant Club", rarity: 3, level: 57 },
    { name: "Reinforced Mace", rarity: 2, level: 63 },
    { name: "Devil Star", rarity: 3, level: 70 },
    { name: "Scourge", rarity: 3, level: 66 },
    { name: "Legendary Mallet", rarity: 3, level: 66 },
    { name: "Ogre Maul", rarity: 3, level: 69 },
    { name: "Thunder Maul", rarity: 2, level: 68 },
    { name: "Falcata", rarity: 1, level: 56 },
    { name: "Ataghan", rarity: 1, level: 61 },
    { name: "Elegant Blade", rarity: 1, level: 63 },
    { name: "Hydra Edge", rarity: 3, level: 69 },
    { name: "Phase Blade", rarity: 2, level: 66 },
    { name: "Conquest Sword", rarity: 2, level: 66 },
    { name: "Cryptic Sword", rarity: 2, level: 66 },
    { name: "Mythical Sword", rarity: 2, level: 66 },
    { name: "Legend Sword", rarity: 3, level: 59 },
    { name: "Highland Blade", rarity: 2, level: 66 },
    { name: "Balrog Blade", rarity: 2, level: 66 },
    { name: "Champion Sword", rarity: 2, level: 66 },
    { name: "Colossal Sword", rarity: 3, level: 66 },
    { name: "Colossus Blade", rarity: 3, level: 66 },
    { name: "Bone Knife", rarity: 2, level: 58 },
    { name: "Mithral Point", rarity: 2, level: 56 },
    { name: "Fanged Knife", rarity: 2, level: 66 },
    { name: "Legend Spike", rarity: 3, level: 66 },
    { name: "Flying Knife", rarity: 2, level: 56 },
    { name: "Flying Axe", rarity: 2, level: 56 },
    { name: "Winged Knife", rarity: 2, level: 77 },
    { name: "Winged Axe", rarity: 2, level: 66 },
    { name: "Hyperion Javelin", rarity: 2, level: 54 },
    { name: "Stygian Pilum", rarity: 2, level: 62 },
    { name: "Balrog Spear", rarity: 3, level: 71 },
    { name: "Ghost Glaive", rarity: 2, level: 79 },
    { name: "Winged Harpoon", rarity: 2, level: 66 },
    { name: "Hyperion Spear", rarity: 2, level: 58 },
    { name: "Stygian Pike", rarity: 2, level: 66 },
    { name: "Mancatcher", rarity: 2, level: 65 },
    { name: "Ghost Spear", rarity: 2, level: 66 },
    { name: "War Pike", rarity: 2, level: 66 },
    { name: "Ogre Axe", rarity: 2, level: 60 },
    { name: "Colossus Voulge", rarity: 2, level: 64 },
    { name: "Thresher", rarity: 2, level: 66 },
    { name: "Cryptic Axe", rarity: 2, level: 66 },
    { name: "Great Poleaxe", rarity: 2, level: 66 },
    { name: "Giant Thresher", rarity: 3, level: 66 },
    { name: "Walking Stick", rarity: 2, level: 58 },
    { name: "Stalagmite", rarity: 2, level: 66 },
    { name: "Elder Staff", rarity: 2, level: 74 },
    { name: "Shillelagh", rarity: 1, level: 66 },
    { name: "Archon Staff", rarity: 1, level: 66 },
    { name: "Spider Bow", rarity: 1, level: 55 },
    { name: "Blade Bow", rarity: 1, level: 60 },
    { name: "Shadow Bow", rarity: 1, level: 63 },
    { name: "Great Bow", rarity: 1, level: 68 },
    { name: "Diamond Bow", rarity: 1, level: 65 },
    { name: "Crusader Bow", rarity: 1, level: 65 },
    { name: "Ward Bow", rarity: 1, level: 66 },
    { name: "Hydra Bow", rarity: 1, level: 66 },
    { name: "Pellet Bow", rarity: 1, level: 57 },
    { name: "Gorgon Crossbow", rarity: 1, level: 67 },
    { name: "Colossus Crossbow", rarity: 1, level: 75 },
    { name: "Demon Crossbow", rarity: 1, level: 65 },
    { name: "Eagle Orb", rarity: 2, level: 1 },
    { name: "Sacred Globe", rarity: 2, level: 8 },
    { name: "Smoked Sphere", rarity: 2, level: 12 },
    { name: "Clasped Orb", rarity: 2, level: 17 },
    { name: "Jared's Stone", rarity: 3, level: 24 },
    { name: "Stag Bow", rarity: 2, level: 18 },
    { name: "Reflex Bow", rarity: 2, level: 27 },
    { name: "Maiden Spear", rarity: 3, level: 18 },
    { name: "Maiden Pike", rarity: 3, level: 27 },
    { name: "Maiden Javelin", rarity: 3, level: 23 },
    { name: "Glowing Orb", rarity: 2, level: 32 },
    { name: "Crystalline Globe", rarity: 3, level: 37 },
    { name: "Cloudy Sphere", rarity: 3, level: 41 },
    { name: "Sparkling Ball", rarity: 2, level: 46 },
    { name: "Swirling Crystal", rarity: 3, level: 50 },
    { name: "Ashwood Bow", rarity: 2, level: 39 },
    { name: "Ceremonial Bow", rarity: 2, level: 47 },
    { name: "Ceremonial Spear", rarity: 2, level: 43 },
    { name: "Ceremonial Pike", rarity: 2, level: 51 },
    { name: "Ceremonial Javelin", rarity: 2, level: 35 },
    { name: "Heavenly Stone", rarity: 2, level: 59 },
    { name: "Eldritch Orb", rarity: 2, level: 67 },
    { name: "Demon Heart", rarity: 2, level: 66 },
    { name: "Vortex Orb", rarity: 3, level: 66 },
    { name: "Dimensional Shard", rarity: 3, level: 66 },
    { name: "Matriarchal Bow", rarity: 3, level: 53 },
    { name: "Grand Matron Bow", rarity: 3, level: 78 },
    { name: "Matriarchal Spear", rarity: 3, level: 61 },
    { name: "Matriarchal Pike", rarity: 3, level: 66 },
    { name: "Matriarchal Javelin", rarity: 3, level: 65 },
  ];

  function findRarityByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.rarity : null; // Return null if not found
  }
  function findLevelByName(name) {
    const rec1 = NameAndRarity.find(rc1 => rc1.name === name);
    return rec1 ? rec1.level : null; // Return null if not found
  }

  const weaponsFilename = 'global\\excel\\weapons.txt';
  const weapons = D2RMM.readTsv(weaponsFilename);
  weapons.rows.forEach((row) => {
    const theName = row['name'];
    const rarity = findRarityByName(theName);
    const level = findLevelByName(theName);
    if (rarity !== null) {
      row.rarity = rarity;
    }
    if (level !== null) {
      row.level = level;
    }
  });
  D2RMM.writeTsv(weaponsFilename, weapons);
}

// StackableGems
{
  const SINGLE_ITEM_CODE = 'gem';
  const STACK_ITEM_CODE = 'sgem';

  const ITEM_TYPES = [];
  function converItemTypeToStackItemType(itemtype) {
    if (itemtype != null && ITEM_TYPES.indexOf(itemtype) !== -1) {
      // original idea to use "z" as prefix ran into issues due to "zlb" already being taken
      // in favor of backwards compatibility, only changing this one conflict
      const prefix = itemtype === 'glb' ? 'q' : 'z';
      return `${prefix}${itemtype.slice(1)}`;
    }
    return itemtype;
  }

  const miscFilenames = [];

  const itemsFilename = 'hd\\items\\items.json';
  const items = D2RMM.readJson(itemsFilename);
  const newItems = [...items];
  for (const index in items) {
    const item = items[index];
    for (const itemtype in item) {
      const asset = item[itemtype].asset;
      if (
        asset.startsWith(`${SINGLE_ITEM_CODE}/`) &&
        !asset.endsWith('_stack') &&
        itemtype !== 'jew' // exclude jewels
      ) {
        ITEM_TYPES.push(itemtype);
        const itemtypeStack = converItemTypeToStackItemType(itemtype);
        newItems.push({ [itemtypeStack]: { asset: `${asset}_stack` } });
        miscFilenames.push(asset.replace(`${SINGLE_ITEM_CODE}/`, ''));
      }
    }
  }
  D2RMM.writeJson(itemsFilename, newItems);

  const miscDirFilename = `hd\\items\\misc\\${SINGLE_ITEM_CODE}\\`;
  for (const index in miscFilenames) {
    const miscFilename = `${miscDirFilename + miscFilenames[index]}.json`;
    const miscStackFilename = `${miscDirFilename + miscFilenames[index]
      }_stack.json`;
    const miscStack = D2RMM.readJson(miscFilename);
    D2RMM.writeJson(miscStackFilename, miscStack);
  }

  const itemtypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemtypesFilename);
  itemtypes.rows.forEach((itemtype) => {
    if (itemtype.Code === SINGLE_ITEM_CODE) {
      itemtypes.rows.push({
        ...itemtype,
        ItemType: `${itemtype.ItemType} Stack`,
        Code: STACK_ITEM_CODE,
        Equiv1: 'misc',
        AutoStack: 1,
      });
    }
  });
  D2RMM.writeTsv(itemtypesFilename, itemtypes);

  if (config.default) {
    const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
    const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
    treasureclassex.rows.forEach((treasureclass) => {
      treasureclass.Item1 = converItemTypeToStackItemType(treasureclass.Item1);
      treasureclass.Item2 = converItemTypeToStackItemType(treasureclass.Item2);
      treasureclass.Item3 = converItemTypeToStackItemType(treasureclass.Item3);
      treasureclass.Item4 = converItemTypeToStackItemType(treasureclass.Item4);
      treasureclass.Item5 = converItemTypeToStackItemType(treasureclass.Item5);
      treasureclass.Item6 = converItemTypeToStackItemType(treasureclass.Item6);
      treasureclass.Item7 = converItemTypeToStackItemType(treasureclass.Item7);
    });
    D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
  }

  const miscFilename = 'global\\excel\\misc.txt';
  const misc = D2RMM.readTsv(miscFilename);
  misc.rows.forEach((item) => {
    if (ITEM_TYPES.indexOf(item.code) !== -1) {
      const itemStack = {
        ...item,
        name: `${item.name} Stack`,
        compactsave: 0,
        type: STACK_ITEM_CODE,
        code: converItemTypeToStackItemType(item.code),
        stackable: 1,
        minstack: 1,
        maxstack: config.maxStack,
        spawnstack: 1,
        spelldesc: 2,
        spelldescstr: 'StackableGem',
        spelldesccolor: 0,
      };
      delete itemStack.type2;
      misc.rows.push(itemStack);
      item.spawnable = 0;
    }
  });
  D2RMM.writeTsv(miscFilename, misc);

  const itemNamesFilename = 'local\\lng\\strings\\item-names.json';
  const itemNames = D2RMM.readJson(itemNamesFilename);
  ITEM_TYPES.forEach((itemtype) => {
    const itemName = itemNames.find(({ Key }) => Key === itemtype);
    if (itemName != null) {
      const stacktype = converItemTypeToStackItemType(itemtype);
      itemNames.push({
        ...itemName,
        id: D2RMM.getNextStringID(),
        Key: stacktype,
      });
    }
  });
  D2RMM.writeJson(itemNamesFilename, itemNames);

  const itemModifiersFilename = 'local\\lng\\strings\\item-modifiers.json';
  const itemModifiers = D2RMM.readJson(itemModifiersFilename);
  itemModifiers.push({
    id: D2RMM.getNextStringID(),
    Key: 'StackableGem',
    enUS: 'Can be transmuted into a usable gem',
    zhTW: '可以转化为可用的宝石',
    deDE: 'Kann in einen nutzbaren Edelstein umgewandelt werden',
    esES: 'Se puede transmutar en una gema utilizable',
    frFR: 'Peut être transmuté en gemme utilisable',
    itIT: 'Può essere trasmutato in una gemma utilizzabile',
    koKR: '사용 가능한 보석으로 변환할 수 있습니다',
    plPL: 'Może zostać przekształcony w użyteczny klejnot',
    esMX: 'Se puede transmutar en una gema utilizable',
    jaJP: '使用可能な宝石に変換することができます',
    ptBR: 'Pode ser transmutado em uma gema utilizável',
    ruRU: 'Может быть превращен в полезный драгоценный камень',
    zhCN: '可以转化为可用的宝石',
  });
  D2RMM.writeJson(itemModifiersFilename, itemModifiers);

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
    const itemtype = ITEM_TYPES[i];
    const stacktype = converItemTypeToStackItemType(itemtype);
    // convert from single to stack
    cubemain.rows.push({
      description: `${itemtype} -> ${stacktype}`,
      enabled: 1,
      version: 100,
      numinputs: 1,
      'input 1': itemtype,
      output: stacktype,
      '*eol\r': 0,
    });
    // convert from stack to single
    cubemain.rows.push({
      description: `${stacktype} -> ${itemtype}`,
      enabled: 1,
      version: 100,
      op: 18, // skip recipe if item's Stat.Accr(param) != value
      param: 70, // quantity (itemstatcost.txt)
      value: 1, // only execute rule if quantity == 1
      numinputs: 1,
      'input 1': stacktype,
      output: itemtype,
      '*eol\r': 0,
    });
  }

  // this is behind a config option because it's *a lot* of recipes...
  if (config.convertWhenDestacking) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      for (let j = 2; j <= config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description: `Stack of ${j} ${itemtype} -> Stack of ${j - 1
            } and Stack of 1`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 1,
          'input 1': stacktype,
          output: `"${stacktype},qty=${j - 1}"`,
          'output b': `"${itemtype},qty=1"`,
          '*eol\r': 0,
        });
      }
    }
  }
  // if another mod already added destacking, don't add it again
  else if (
    cubemain.rows.find(
      (row) => row.description === 'Stack of 2 -> Stack of 1 and Stack of 1'
    ) == null
  ) {
    for (let i = 2; i <= config.maxStack; i = i + 1) {
      cubemain.rows.push({
        description: `Stack of ${i} -> Stack of ${i - 1} and Stack of 1`,
        enabled: 1,
        version: 0,
        op: 18, // skip recipe if item's Stat.Accr(param) != value
        param: 70, // quantity (itemstatcost.txt)
        value: i, // only execute rule if quantity == i
        numinputs: 1,
        'input 1': 'misc',
        output: `"usetype,qty=${i - 1}"`,
        'output b': `"usetype,qty=1"`,
        '*eol\r': 0,
      });
    }
  }

  if (config.bulkUpgrade) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      // no upgrade for perfect gems
      if ((i + 1) % 5 == 0) {
        continue;
      }
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      const upgradedItemtype = ITEM_TYPES[i + 1];
      const upgradedStacktype = converItemTypeToStackItemType(upgradedItemtype);
      for (let j = 30; j < config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description:
            `Stack of ${j} ${itemtype} + 1 id scroll -> Stack` +
            ` of 10 ${upgradedItemtype} + Stack of ${j - 30} ${itemtype} + 1` +
            ` id scroll`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 2,
          'input 1': stacktype,
          'input 2': 'isc',
          output: `"${upgradedStacktype},qty=10"`,
          'output b': j == 30 ? null : `"${stacktype},qty=${j - 30}"`,
          'output c': 'isc',
          '*eol\r': 0,
        });
      }
    }
  }
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// StackableRunes
{
  const SINGLE_ITEM_CODE = 'rune';
  const STACK_ITEM_CODE = 'runs';

  const ITEM_TYPES = [];
  function converItemTypeToStackItemType(itemtype) {
    if (itemtype != null && ITEM_TYPES.indexOf(itemtype) !== -1) {
      return itemtype.replace(/^r/, 's');
    }
    return itemtype;
  }

  const miscFilenames = [];

  const itemsFilename = 'hd\\items\\items.json';
  const items = D2RMM.readJson(itemsFilename);
  const newItems = [...items];
  for (const index in items) {
    const item = items[index];
    for (const itemtype in item) {
      const asset = item[itemtype].asset;
      if (asset.startsWith(`${SINGLE_ITEM_CODE}/`) && !asset.endsWith('_stack')) {
        ITEM_TYPES.push(itemtype);
        const itemtypeStack = converItemTypeToStackItemType(itemtype);
        newItems.push({ [itemtypeStack]: { asset: `${asset}_stack` } });
        miscFilenames.push(asset.replace(`${SINGLE_ITEM_CODE}/`, ''));
      }
    }
  }
  D2RMM.writeJson(itemsFilename, newItems);

  const miscDirFilename = `hd\\items\\misc\\${SINGLE_ITEM_CODE}\\`;
  for (const index in miscFilenames) {
    const miscFilename = `${miscDirFilename + miscFilenames[index]}.json`;
    const miscStackFilename = `${miscDirFilename + miscFilenames[index]
      }_stack.json`;
    const miscStack = D2RMM.readJson(miscFilename);
    D2RMM.writeJson(miscStackFilename, miscStack);
  }

  const itemtypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemtypesFilename);
  itemtypes.rows.forEach((itemtype) => {
    if (itemtype.Code === SINGLE_ITEM_CODE) {
      itemtypes.rows.push({
        ...itemtype,
        ItemType: `${itemtype.ItemType} Stack`,
        Code: STACK_ITEM_CODE,
        Equiv1: 'misc',
        AutoStack: 1,
      });
    }
  });
  D2RMM.writeTsv(itemtypesFilename, itemtypes);

  if (config.default) {
    const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
    const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
    treasureclassex.rows.forEach((treasureclass) => {
      treasureclass.Item1 = converItemTypeToStackItemType(treasureclass.Item1);
      treasureclass.Item2 = converItemTypeToStackItemType(treasureclass.Item2);
    });
    D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
  }

  const miscFilename = 'global\\excel\\misc.txt';
  const misc = D2RMM.readTsv(miscFilename);
  misc.rows.forEach((item) => {
    if (ITEM_TYPES.indexOf(item.code) !== -1) {
      misc.rows.push({
        ...item,
        name: `${item.name} Stack`,
        compactsave: 0,
        type: STACK_ITEM_CODE,
        code: converItemTypeToStackItemType(item.code),
        stackable: 1,
        minstack: 1,
        maxstack: config.maxStack,
        spawnstack: 1,
        spelldesc: 2,
        spelldescstr: 'StackableRune',
        spelldesccolor: 0,
      });
      item.spawnable = 0;
    }
  });
  D2RMM.writeTsv(miscFilename, misc);

  const itemNamesFilename = 'local\\lng\\strings\\item-names.json';
  const itemNames = D2RMM.readJson(itemNamesFilename);
  ITEM_TYPES.forEach((itemtype) => {
    const itemName = itemNames.find(({ Key }) => Key === itemtype);
    if (itemName != null) {
      const stacktype = converItemTypeToStackItemType(itemtype);
      itemNames.push({
        ...itemName,
        id: D2RMM.getNextStringID(),
        Key: stacktype,
      });
    }
  });
  D2RMM.writeJson(itemNamesFilename, itemNames);

  const itemModifiersFilename = 'local\\lng\\strings\\item-modifiers.json';
  const itemModifiers = D2RMM.readJson(itemModifiersFilename);
  itemModifiers.push({
    id: D2RMM.getNextStringID(),
    Key: 'StackableRune',
    enUS: 'Can be transmuted into a usable rune',
    zhTW: '可以转化为可用的符文',
    deDE: 'Kann in eine nutzbare Rune umgewandelt werden',
    esES: 'Se puede transmutar en una runa utilizable',
    frFR: 'Peut être transmuté en une rune utilisable',
    itIT: 'Può essere trasmutato in una runa utilizzabile',
    koKR: '사용 가능한 룬으로 변환할 수 있습니다',
    plPL: 'Może zostać przemieniony w użyteczną runę',
    esMX: 'Se puede transmutar en una runa utilizable',
    jaJP: '使用可能なルーンに変換できます',
    ptBR: 'Pode ser transmutado em uma runa utilizável',
    ruRU: 'Может быть преобразован в полезную руну',
    zhCN: '可以转化为可用的符文',
  });
  D2RMM.writeJson(itemModifiersFilename, itemModifiers);

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
    const itemtype = ITEM_TYPES[i];
    const stacktype = converItemTypeToStackItemType(itemtype);
    // convert from single to stack
    cubemain.rows.push({
      description: `${itemtype} -> ${stacktype}`,
      enabled: 1,
      version: 100,
      numinputs: 1,
      'input 1': itemtype,
      output: stacktype,
      '*eol\r': 0,
    });
    // convert from stack to single
    cubemain.rows.push({
      description: `${stacktype} -> ${itemtype}`,
      enabled: 1,
      version: 100,
      op: 18, // skip recipe if item's Stat.Accr(param) != value
      param: 70, // quantity (itemstatcost.txt)
      value: 1, // only execute rule if quantity == 1
      numinputs: 1,
      'input 1': stacktype,
      output: itemtype,
      '*eol\r': 0,
    });
  }

  // this is behind a config option because it's *a lot* of recipes...
  if (config.convertWhenDestacking) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      for (let j = 2; j <= config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description: `Stack of ${j} ${itemtype} -> Stack of ${j - 1
            } and Stack of 1`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 1,
          'input 1': stacktype,
          output: `"${stacktype},qty=${j - 1}"`,
          'output b': `"${itemtype},qty=1"`,
          '*eol\r': 0,
        });
      }
    }
  }
  // if another mod already added destacking, don't add it again
  else if (
    cubemain.rows.find(
      (row) => row.description === 'Stack of 2 -> Stack of 1 and Stack of 1'
    ) == null
  ) {
    for (let i = 2; i <= config.maxStack; i = i + 1) {
      cubemain.rows.push({
        description: `Stack of ${i} -> Stack of ${i - 1} and Stack of 1`,
        enabled: 1,
        version: 0,
        op: 18, // skip recipe if item's Stat.Accr(param) != value
        param: 70, // quantity (itemstatcost.txt)
        value: i, // only execute rule if quantity == i
        numinputs: 1,
        'input 1': 'misc',
        output: `"usetype,qty=${i - 1}"`,
        'output b': `"usetype,qty=1"`,
        '*eol\r': 0,
      });
    }
  }
  D2RMM.writeTsv(cubemainFilename, cubemain);

  // D2R colors runes as orange by default, but it seems to be based on item type
  // rather than localization strings so it does not apply to the stacked versions
  // we update the localization file to manually color the names of runes here
  // so that it will also apply to the stacked versions of the runes
  const itemRunesFilename = 'local\\lng\\strings\\item-runes.json';
  const itemRunes = D2RMM.readJson(itemRunesFilename);
  itemRunes.forEach((item) => {
    const itemtype = item.Key;
    if (itemtype.match(/^r[0-9]{2}$/) != null) {
      // update all localizations
      for (const key in item) {
        if (key !== 'id' && key !== 'Key') {
          if (item[key].indexOf('ÿc') !== -1) {
            // if the rune is already colored by something else (e.g. another mod)
            // then respect that change and don't override it
            continue;
          }
          // no idea what this is, but color codes before [fs] don't work
          const [, prefix = '', value] = item[key].match(/^(\[fs\])?(.*)$/) ?? [
            '',
            '',
            item[key],
          ];
          item[key] = `${prefix}ÿc8${value}`;
        }
      }
    }
  });
  D2RMM.writeJson(itemRunesFilename, itemRunes);
}

// MrLamaSc character startup
{
  const charstatsFilename = 'global\\excel\\charstats.txt';
  const charstats = D2RMM.readTsv(charstatsFilename);
  charstats.rows.forEach((row) => {
    const id = row.class;
    if (id === 'Sorceress') {
      row.StartSkill = 'charged bolt';
    }
    if (id === 'Amazon' || id === 'Assassin' || id === 'Barbarian' || id === 'Druid' || id === 'Paladin') {
      row.item2 = 'hp1';
      row.item2loc = '';
      row.item2count = 4;
      row.item3 = 'tsc';
      row.item3count = 1;
      row.item4 = 'isc';
      row.item5 = 0;
      row.item5count = 0;
    }
  });
  D2RMM.writeTsv(charstatsFilename, charstats);
}

// Move Cain
{
  if (config["act5_options"] == "stash") {
    D2RMM.copyFile(
      'act5_cain_stash\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_stash\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }

  if (config["act5_options"] == "malah") {
    D2RMM.copyFile(
      'act5_cain_malah\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_malah\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }

  if (config["act5_options"] == "larzuk") {
    D2RMM.copyFile(
      'act5_cain_larzuk\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_larzuk\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }
}

// Copy files from hd to hd
{
  D2RMM.copyFile(
    'hd', // <mod folder>\hd
    'hd', // <diablo 2 folder>\mods\<modname>\<modname>.mpq\data\hd
    true // overwrite any conflicts
  );
}

// Modify Stash
{
  // modify the stash save file to make sure it has 8 tab pages
  if (config.isExtraTabsEnabled) {
    function getStashTabBinary(vers) {
      // prettier-ignore
      return [
        // extracted from first 68 bytes of an empty stash save file (SharedStashSoftCoreV2.d2i) of D2R v1.6.80273
        // the 9th byte seems to be related to the version of the game (e.g. 0x61, 0x62, 0x63) and a stash save
        // that uses a *newer* version than the currently running version of the game will fail to load
        0x55, 0xAA, 0x55, 0xAA, 0x01, 0x00, 0x00, 0x00, vers, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
        0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
        0x4A, 0x4D, 0x00, 0x00                                                                          //  4 bytes
      ];
    }

    function indexOf(haystack, needle, startIndex = 0) {
      let match = 0;
      const index = haystack.findIndex((value, index) => {
        if (index < startIndex) {
          return false;
        }
        // null matches any value
        if (needle[match] == null || value === needle[match]) {
          match++;
        } else {
          match = 0;
        }
        if (match === needle.length) {
          return true;
        }
        return false;
      });
      return index === -1 ? -1 : index - needle.length + 1;
    }

    function getStashTabStartIndices(stashData) {
      const stashTabPrefix = getStashTabBinary(null).slice(0, 10);
      const stashTabStartIndices = [];
      let index = -1;
      while (true) {
        index = indexOf(stashData, stashTabPrefix, index + 1);
        if (index !== -1) {
          stashTabStartIndices.push(index);
        } else {
          break;
        }
      }
      return stashTabStartIndices;
    }

    const results = {};
    function modSaveFile(filename) {
      const stashData = D2RMM.readSaveFile(filename);
      if (stashData == null) {
        console.debug(`Skipped ${filename} because the file was not found.`);
        results[filename] = false;
        return;
      }
      // backup existing stash tab data if it doesn't exist
      const stashDataBackup = D2RMM.readSaveFile(`${filename}.bak`);
      if (stashDataBackup == null) {
        D2RMM.writeSaveFile(`${filename}.bak`, stashData);
      }
      // find number of times prefix appears in stashData
      const stashTabIndices = getStashTabStartIndices(stashData);
      // find latest version code used by the save file
      const versionCode = Math.max(
        ...stashTabIndices.map((index) => stashData[index + 8])
      );
      // sanitize the data (each save files should have 3-7 shared tabs)
      const existingTabsCount = Math.max(3, Math.min(7, stashTabIndices.length));
      const tabsToAdd = 7 - existingTabsCount;
      // don't modify the save file if it doesn't need it
      if (tabsToAdd > 0) {
        D2RMM.writeSaveFile(
          filename,
          [].concat.apply(
            stashData,
            new Array(tabsToAdd).fill(getStashTabBinary(versionCode))
          )
        );
        console.debug(
          `Added ${tabsToAdd} additional shared stash tabs to ${filename} using version code 0x${versionCode.toString(
            16
          )}.`
        );
      } else {
        console.debug(
          `Skipped ${filename} because it already has 7 shared stash tabs.`
        );
      }
      results[filename] = true;
    }

    const SOFTCORE_SAVE_FILE = 'SharedStashSoftCoreV2.d2i';
    const HARDCORE_SAVE_FILE = 'SharedStashHardCoreV2.d2i';

    modSaveFile(SOFTCORE_SAVE_FILE);
    modSaveFile(HARDCORE_SAVE_FILE);

    if (!results[SOFTCORE_SAVE_FILE] && !results[HARDCORE_SAVE_FILE]) {
      console.warn(
        `Unable to enable additional shared stash tabs. Neither ${SOFTCORE_SAVE_FILE} nor ${HARDCORE_SAVE_FILE} were found.`
      );
    }
  }
}
