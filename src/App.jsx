import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { useSeededCollection, saveItem, deleteItem, replaceCollection } from "./lib/firestoreData";
import { loadGoogleMaps, geocodeAddress } from "./lib/googleMaps";
import {
  Bus, User, MapPin, CreditCard, HeartPulse, ChevronRight, Milestone,
  Calendar, Navigation, Search, Plus, Pencil, X, Map as MapIcon, List, Trash2,
  Home, Sun, Moon, Users, Settings, UserPlus, LogOut, Building2, Menu,
  Download, Printer, FileSpreadsheet, Upload, Lock, Pin, PinOff, Route,
} from "lucide-react";

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
`;

const LOGO_URL = "/logo.png";

const INK = "#23342E";
const MUTED = "#8B978F";
const GREEN = "#3F6C51";
const OCHRE = "#C97B3D";

const INSTITUCIONES_INICIAL = [
  { id: "sedeChile", nombre: "Sede Chile", direccion: "Completar dirección", lat: -26.536884, lng: -59.333343 },
  { id: "sedeIrigoyen", nombre: "Sede Irigoyen", direccion: "Completar dirección", lat: -26.536884, lng: -59.333343 },
  { id: "sedeUruguay", nombre: "Sede Uruguay", direccion: "Uruguay 1129, General José de San Martín, Chaco", lat: -26.536884, lng: -59.333343 },
  { id: "sedeUnidas", nombre: "Sede Unidas", direccion: "Completar dirección", lat: -26.536884, lng: -59.333343 },
];

const PRESTACIONES_INICIAL = [
  { id: "p1", nombre: "SAIE", institucionId: "sedeChile" },
  { id: "p2", nombre: "SET", institucionId: "sedeChile" },
  { id: "p3", nombre: "CET", institucionId: "sedeIrigoyen" },
  { id: "p4", nombre: "Centro de Día", institucionId: "sedeUruguay" },
  { id: "p5", nombre: "Hogar Permanente", institucionId: "sedeUnidas" },
];

const LOCALIDADES_INICIAL = [
  { id: "loc1", nombre: "Bermejo", provincia: "Chaco" },
  { id: "loc2", nombre: "Capitán Solari", provincia: "Chaco" },
  { id: "loc3", nombre: "Ciervo", provincia: "Chaco" },
  { id: "loc4", nombre: "Colonia Elisa", provincia: "Chaco" },
  { id: "loc5", nombre: "Colonias Unidas", provincia: "Chaco" },
  { id: "loc6", nombre: "El Colorado", provincia: "Chaco" },
  { id: "loc7", nombre: "Elisa", provincia: "Chaco" },
  { id: "loc8", nombre: "General Vedia", provincia: "Chaco" },
  { id: "loc9", nombre: "General Vedia (Campo)", provincia: "Chaco" },
  { id: "loc10", nombre: "Gral Jose De San Martin", provincia: "Chaco" },
  { id: "loc11", nombre: "Gral San Martin", provincia: "Chaco" },
  { id: "loc12", nombre: "Gral. De San Martin", provincia: "Chaco" },
  { id: "loc13", nombre: "Gral.Jose De San Martin", provincia: "Chaco" },
  { id: "loc14", nombre: "Gral.San Martin", provincia: "Chaco" },
  { id: "loc15", nombre: "La Escondida", provincia: "Chaco" },
  { id: "loc16", nombre: "La Laonesa", provincia: "Chaco" },
  { id: "loc17", nombre: "La Leonesa", provincia: "Chaco" },
  { id: "loc18", nombre: "La Verde", provincia: "Chaco" },
  { id: "loc19", nombre: "Lapachito", provincia: "Chaco" },
  { id: "loc20", nombre: "Las Palmas", provincia: "Chaco" },
  { id: "loc21", nombre: "Presidencia Roca", provincia: "Chaco" },
  { id: "loc22", nombre: "Villa 213", provincia: "Chaco" },
];

const OBRAS_SOCIALES_INICIAL = [
  { id: "os1", nombre: "PAMI" },
  { id: "os2", nombre: "IOMA" },
  { id: "os3", nombre: "OSDE" },
  { id: "os4", nombre: "OSEP" },
  { id: "os5", nombre: "Particular" },
];

const RECORRIDOS_INICIAL = [
  {
    id: "r1", nombre: "Roca", localidad: "Presidencia Roca",
    chofer: "MALDONADO OMAR", auxiliar: "ZACARIAS GUSTAVO", vehiculo: "", patente: "AH 301 TG",
    kmIda: 62.4,
    chicos: [
      { nombre: "Aguirre Andrea Adalis", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A  VIERNES", institucionId: "sedeUruguay", kmDesde: 60.2, lat: -26.144265, lng: -59.599022 },
      { nombre: "Aguirre Pedro De La Cruz", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.9, lat: -26.140958, lng: -59.599093 },
      { nombre: "Alegre Maria Magdalena", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.9, lat: -26.140264, lng: -59.599505 },
      { nombre: "Barbosa Santa", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 59.7, lat: -26.14676, lng: -59.594444 },
      { nombre: "Borda Lautaro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 59.5, lat: -26.144636, lng: -59.592023 },
      { nombre: "Cardozo Carina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 59.5, lat: -26.14424, lng: -59.592668 },
      { nombre: "Chavez Agustina", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "MIERCOLES Y VIERNES", institucionId: "sedeChile", kmDesde: 60.4, lat: -26.145485, lng: -59.591685 },
      { nombre: "Fioroni Nilda", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A  VIERNES", institucionId: "sedeUruguay", kmDesde: 60.3, lat: -26.146774, lng: -59.599333 },
      { nombre: "Gimenez Cristian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.7, lat: -26.138062, lng: -59.5974 },
      { nombre: "Gimenez Franco", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.3, lat: -26.146585, lng: -59.599248 },
      { nombre: "Gomez Prisila", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.3, lat: -26.136903, lng: -59.591975 },
      { nombre: "Paredes Alejandro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.4, lat: -26.136164, lng: -59.591724 },
      { nombre: "Paredes Ramona Beatris", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.4, lat: -26.136131, lng: -59.591984 },
      { nombre: "Paredes Ramona Ofelia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 60.5, lat: -26.140973, lng: -59.597803 },
      { nombre: "Parra Maria Jose", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 59.8, lat: -26.145479, lng: -59.596742 },
      { nombre: "Silva Gimez Lourdes", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Presidencia Roca", dias: "MIERCOLES Y VIERNES", institucionId: "sedeChile", kmDesde: 62.4, lat: -26.138916, lng: -59.599793 },
    ],
  },
  {
    id: "r2", nombre: "San Martin 1", localidad: "Gral. De San Martin",
    chofer: "BLUME OSVALDO (IDA Y VUELTA)", auxiliar: "RIVERO WALTER (IDA Y VUELTA)", vehiculo: "", patente: "AB 128 IW",
    kmIda: 2.8,
    chicos: [
      { nombre: "Balmaceda Waldo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.544498, lng: -59.322743 },
      { nombre: "Cueva Marcelo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.0, lat: -26.544309, lng: -59.326068 },
      { nombre: "Dure Julio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.541463, lng: -59.326694 },
      { nombre: "Fernandez Silvia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.0, lat: -26.530471, lng: -59.334037 },
      { nombre: "Lopez Jonathan", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.525074, lng: -59.331717 },
      { nombre: "Lopez Juan", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.5, lat: -26.544688, lng: -59.322552 },
      { nombre: "Lopez Lucas", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.7, lat: -26.546281, lng: -59.321371 },
      { nombre: "Molina Aldana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.541661, lng: -59.326749 },
      { nombre: "Moreira Milagros", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 0.8, lat: -26.532109, lng: -59.334133 },
      { nombre: "Moron Lautaro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.6, lat: -26.549386, lng: -59.32281 },
      { nombre: "Peñalver Eduardo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.524279, lng: -59.330942 },
      { nombre: "Ramirez Mathias", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.523777, lng: -59.332013 },
      { nombre: "Samana Benigno", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.543831, lng: -59.322834 },
    ],
  },
  {
    id: "r3", nombre: "San Martin 2", localidad: "Gral. De San Martin",
    chofer: "BLUME OSVALDO (IDA Y VUELTA)", auxiliar: "MACIEL NELSON-SANCHEZ CARLA  (IDA)  ACOSTA JONATAN-ZACARIAS SILVIA (VUELTA)", vehiculo: "", patente: "AB 128 IW",
    kmIda: 4.0,
    chicos: [
      { nombre: "Aguirre Nicolas", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 0.7, lat: -26.54028, lng: -59.334211 },
      { nombre: "Arriola Fabian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.544865, lng: -59.331482 },
      { nombre: "Casco Vicente", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.2, lat: -26.542233, lng: -59.331468 },
      { nombre: "Debes Micaela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.547394, lng: -59.331978 },
      { nombre: "Duarte Maria", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.3, lat: -26.541187, lng: -59.331503 },
      { nombre: "Maidana Malvina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.545471, lng: -59.330001 },
      { nombre: "Monzon Lucia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.547392, lng: -59.332066 },
      { nombre: "Pereyra Gaston", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 4.0, lat: -26.555601, lng: -59.321664 },
      { nombre: "Silguero Agustin", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.9, lat: -26.55455, lng: -59.335034 },
      { nombre: "Talavera Jessica", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.548042, lng: -59.333527 },
      { nombre: "Torres Emilce", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.0, lat: -26.541372, lng: -59.333877 },
      { nombre: "Valdez Maximo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.9, lat: -26.5565, lng: -59.330436 },
      { nombre: "Villalva Daniel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.551933, lng: -59.336531 },
      { nombre: "Villalva Dario", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral. De San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.551933, lng: -59.336531 },
    ],
  },
  {
    id: "r4", nombre: "San Martin 3", localidad: "Gral San Martin",
    chofer: "OCAMPO HORACIO", auxiliar: "GALARZA AGUSTIN-GODOY JAVIER  (IDA) PINO RAMON-GODOY MARCELO (VUELTA)", vehiculo: "", patente: "KYK 330",
    kmIda: 3.5,
    chicos: [
      { nombre: "Bravo Braian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.532491, lng: -59.345987 },
      { nombre: "Capello Liliana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LU,MAR,JUE,VIER", institucionId: "sedeUruguay", kmDesde: 1.5, lat: -26.531565, lng: -59.343706 },
      { nombre: "Castro Daniel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.3, lat: -26.52993, lng: -59.339193 },
      { nombre: "Correa Milagros", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.532511, lng: -59.346252 },
      { nombre: "Gomez Bruno", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.529195, lng: -59.351602 },
      { nombre: "Lopez Julio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.7, lat: -26.537017, lng: -59.344531 },
      { nombre: "Lopez Natalia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.5, lat: -26.537649, lng: -59.357948 },
      { nombre: "Paetz Marisa", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 0.5, lat: -26.53338, lng: -59.335939 },
      { nombre: "Palavecino Jesus", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.534853, lng: -59.347992 },
      { nombre: "Ramirez Adrian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 0.4, lat: -26.534863, lng: -59.333142 },
      { nombre: "Ramirez Silvana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.6, lat: -26.530203, lng: -59.352607 },
      { nombre: "Vallejos Martin", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.0, lat: -26.532451, lng: -59.34957 },
    ],
  },
  {
    id: "r5", nombre: "San Martin 4", localidad: "Gral San Martin",
    chofer: "LEANDRO TOFFOLETI", auxiliar: "CARRASCO ALEJANDRO- INSAURRALDE ROCIO(IDA)  PEREYRA JUAN.C-FEMENIA CINTHIA (VUELTA)", vehiculo: "", patente: "OBP 418",
    kmIda: 2.4,
    chicos: [
      { nombre: "Alegre Lisandro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.551423, lng: -59.337496 },
      { nombre: "Aquino Diego", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.551014, lng: -59.339528 },
      { nombre: "Aquino M.Del Carmen", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.551014, lng: -59.339528 },
      { nombre: "Aquino M.Florencia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.551014, lng: -59.339528 },
      { nombre: "Baez Graciela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.4, lat: -26.545433, lng: -59.334644 },
      { nombre: "Costich Sandro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.547545, lng: -59.339064 },
      { nombre: "Filipon Nancy", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.6, lat: -26.546538, lng: -59.33689 },
      { nombre: "Flores Macarena", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.547301, lng: -59.338893 },
      { nombre: "Gomez Carlos", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.0, lat: -26.54909, lng: -59.338863 },
      { nombre: "Insauralde Guadalupe", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.9, lat: -26.547911, lng: -59.33775 },
      { nombre: "Insaurralde Marianela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.6, lat: -26.548986, lng: -59.340238 },
      { nombre: "Lopez Erminda", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.5496, lng: -59.336254 },
      { nombre: "Martinez Arnaldo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.6, lat: -26.54605, lng: -59.338473 },
      { nombre: "Rivarola Angel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.549657, lng: -59.340548 },
      { nombre: "Romero Jessica", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.9, lat: -26.547911, lng: -59.33775 },
      { nombre: "Romero Milagros", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.550194, lng: -59.334959 },
      { nombre: "Sandoval Facundo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.5496, lng: -59.336254 },
      { nombre: "Sandoval Mauro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.5496, lng: -59.336254 },
      { nombre: "Sandoval Sebastian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.5496, lng: -59.336254 },
      { nombre: "Saucedo M.Celeste", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.8, lat: -26.546559, lng: -59.335196 },
      { nombre: "Villanueva Javier", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.6, lat: -26.546299, lng: -59.335902 },
    ],
  },
  {
    id: "r6", nombre: "San Martin 5", localidad: "Gral.San Martin",
    chofer: "OCAMPO HORACIO", auxiliar: "CACERES ELIAS-ROMAN LORENA  (IDA) MEZA PAOLA-RIVERO WALTER (VUELTA)", vehiculo: "", patente: "KYK 330",
    kmIda: 4.2,
    chicos: [
      { nombre: "Luque Orlando", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.0, lat: -26.55236, lng: -59.343315 },
      { nombre: "Cantero Ariel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.548653, lng: -59.347706 },
      { nombre: "Cantero Diego", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.548653, lng: -59.347706 },
      { nombre: "Chiou Angel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.9, lat: -26.545071, lng: -59.352165 },
      { nombre: "Fernandez Marcelo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.548573, lng: -59.3475 },
      { nombre: "Fernandez Ariel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.7, lat: -26.545423, lng: -59.355601 },
      { nombre: "Gaete Eliana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.0, lat: -26.546642, lng: -59.351849 },
      { nombre: "Insaurralde Emilce", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.1, lat: -26.547994, lng: -59.351208 },
      { nombre: "Insaurralde Jessica", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.1, lat: -26.547709, lng: -59.350799 },
      { nombre: "Lopez Mauricio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.552147, lng: -59.344056 },
      { nombre: "Lopez Ricardo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.4, lat: -26.543943, lng: -59.357181 },
      { nombre: "Muñoz Cintia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.8, lat: -26.548626, lng: -59.347804 },
      { nombre: "Osuna Araceli", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 4.2, lat: -26.545092, lng: -59.360375 },
      { nombre: "Salomon Luz", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.2, lat: -26.551674, lng: -59.345408 },
      { nombre: "Villordo Lorenzo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 4.1, lat: -26.546652, lng: -59.357483 },
    ],
  },
  {
    id: "r7", nombre: "San Martin 6", localidad: "Gral.San Martin",
    chofer: "LEANDRO TOFFOLETI", auxiliar: "ALEGRE DALMA-VALLEJOS PAOLA  (IDA) IBARRA VALERIA-BOGDANIC EVELYN  (VUELTA)", vehiculo: "", patente: "OBP 418",
    kmIda: 3.4,
    chicos: [
      { nombre: "Ayala Natalia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.6, lat: -26.543398, lng: -59.349143 },
      { nombre: "De Jesus Milagros", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.2, lat: -26.543359, lng: -59.352921 },
      { nombre: "Encizo Rocio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.5, lat: -26.53771, lng: -59.340844 },
      { nombre: "Farias Celestina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.7, lat: -26.536644, lng: -59.354771 },
      { nombre: "Farias Feliciano", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.6, lat: -26.53665, lng: -59.354011 },
      { nombre: "Galeano Alejandro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.536265, lng: -59.350807 },
      { nombre: "Gomez Marcos", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.53643, lng: -59.350942 },
      { nombre: "Gomez Norma", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 3.4, lat: -26.540907, lng: -59.354237 },
      { nombre: "Jara Yamila", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.9, lat: -26.542364, lng: -59.350529 },
      { nombre: "Ojeda Maria Belen", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.6, lat: -26.542753, lng: -59.350973 },
      { nombre: "Pino Sergio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.7, lat: -26.538496, lng: -59.353085 },
      { nombre: "Romero Zaira", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.4, lat: -26.542717, lng: -59.34541 },
      { nombre: "Sandoval Gladys", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.5, lat: -26.544291, lng: -59.349416 },
      { nombre: "Silvestri Alan", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.9, lat: -26.539054, lng: -59.344783 },
      { nombre: "Sosa Jessica", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.3, lat: -26.54013, lng: -59.349415 },
      { nombre: "Taborda Daiana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 2.9, lat: -26.541695, lng: -59.351772 },
      { nombre: "Turraca Braulio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Gral.San Martin", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 1.5, lat: -26.537392, lng: -59.340736 },
    ],
  },
  {
    id: "r8", nombre: "Unidas", localidad: "Colonias Unidas",
    chofer: "GIMENEZ MARTIN  (IDA) / AGUIRRE CRISTIAN (VUELTA)", auxiliar: "MOYANO HECTOR  (IDA) / ROBLES RODOLFO    (VUELTA) -", vehiculo: "", patente: "PMZ 624",
    kmIda: 42.3,
    chicos: [
      { nombre: "Blanco Diego", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 41.5, lat: -26.698177, lng: -59.625863 },
      { nombre: "Cañete Mercedes", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 41.7, lat: -26.695526, lng: -59.631601 },
      { nombre: "Cañete Gilberto", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 42.3, lat: -26.692848, lng: -59.629956 },
      { nombre: "Chaves Claudio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 41.1, lat: -26.706297, lng: -59.630847 },
      { nombre: "Encina Emanuel P", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 42.2, lat: -26.695319, lng: -59.636839 },
      { nombre: "Encina Hugo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 41.6, lat: -26.697039, lng: -59.626917 },
      { nombre: "Fernandez Angel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 42.2, lat: -26.713022, lng: -59.639384 },
      { nombre: "Fontana Ruben", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 38.2, lat: -26.699926, lng: -59.635564 },
      { nombre: "Maldonado Ruiz Diaz Maximiliano", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "MARTES -JUEVES-VIERNES", institucionId: "sedeChile", kmDesde: 36.5, lat: -26.697693, lng: -59.62637 },
      { nombre: "Marcon  Rodrigo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 37.0, lat: -26.697971, lng: -59.628852 },
      { nombre: "Peralta Silvina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 37.8, lat: -26.693524, lng: -59.626857 },
      { nombre: "Ramirez Oscar", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 37.3, lat: -26.69898, lng: -59.626716 },
      { nombre: "Silva Miguel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 38.0, lat: -26.694267, lng: -59.628965 },
      { nombre: "Sisuela Jose", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 38.3, lat: -26.69713, lng: -59.636272 },
      { nombre: "Zapata Martin", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonias Unidas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 42.0, lat: -26.6941, lng: -59.632354 },
    ],
  },
  {
    id: "r9", nombre: "S.A.I.E", localidad: "Gral.Jose De San Martin",
    chofer: "PRIETO HECTOR  IDA Y VUELTA", auxiliar: "SIN AUXILIUAR", vehiculo: "", patente: "MSF 206",
    kmIda: 4.0,
    chicos: [
      { nombre: "Alegre Giuliano", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 1.9, lat: -26.535071, lng: -59.354507 },
      { nombre: "Frias Dure Benjamin", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Jueves y Viernes", institucionId: "sedeChile", kmDesde: 2.1, lat: -26.542108, lng: -59.33033 },
      { nombre: "Hrycuck Maximiliano", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 1.7, lat: -26.5466, lng: -59.349579 },
      { nombre: "Martin Julieta", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 3.1, lat: -26.546988, lng: -59.322403 },
      { nombre: "Martinez Isaias", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Miercoles y Viernes", institucionId: "sedeChile", kmDesde: 1.8, lat: -26.542816, lng: -59.333023 },
      { nombre: "Mendez Milagros", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes Y Jueves", institucionId: "sedeChile", kmDesde: 2.4, lat: -26.549503, lng: -59.336411 },
      { nombre: "Godoy Nemesio Franco", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Jueves", institucionId: "sedeChile", kmDesde: 3.0, lat: -26.526968, lng: -59.332665 },
      { nombre: "Quiroz Daiana", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 1.5, lat: -26.538517, lng: -59.333897 },
      { nombre: "Rojas Maxima", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 1.5, lat: -26.543423, lng: -59.350273 },
      { nombre: "Sanabria Sofia", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 2.0, lat: -26.54909, lng: -59.338863 },
      { nombre: "Torales Mayco", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Viernes", institucionId: "sedeChile", kmDesde: 1.8, lat: -26.551394, lng: -59.342117 },
      { nombre: "Avalos Enzo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 2.0, lat: -26.547977, lng: -59.333298 },
      { nombre: "Ayala Facundo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 4.0, lat: -26.55716, lng: -59.33021 },
      { nombre: "Ayala Manuel", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 4.0, lat: -26.55716, lng: -59.33021 },
      { nombre: "Barrios Mario", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 1.4, lat: -26.542812, lng: -59.348958 },
      { nombre: "Galarza Valentin", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Viernes", institucionId: "sedeChile", kmDesde: 2.2, lat: -26.551854, lng: -59.340895 },
      { nombre: "Gomez Diaz Juan Luis", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Jueves", institucionId: "sedeChile", kmDesde: 4.0, lat: -26.548954, lng: -59.3236 },
      { nombre: "Gonzales Fernanda", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 2.5, lat: -26.546969, lng: -59.329651 },
      { nombre: "Lescano Velasco Santino", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Miercoles", institucionId: "sedeChile", kmDesde: 1.3, lat: -26.536627, lng: -59.350948 },
      { nombre: "Liva Lucas", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Viernes", institucionId: "sedeChile", kmDesde: 1.8, lat: -26.540882, lng: -59.328764 },
      { nombre: "Lopez Rodrigo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 2.5, lat: -26.552755, lng: -59.346218 },
      { nombre: "Maidana Jonathan", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 3.9, lat: -26.521552, lng: -59.330757 },
      { nombre: "Miranda Franco", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Miercoles", institucionId: "sedeChile", kmDesde: 3.0, lat: -26.526944, lng: -59.332485 },
      { nombre: "Mora Sofia", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Miercoles y jueves", institucionId: "sedeChile", kmDesde: 1.7, lat: -26.541424, lng: -59.351841 },
      { nombre: "Morales Thian", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Miercoles", institucionId: "sedeChile", kmDesde: 4.0, lat: -26.557319, lng: -59.329471 },
      { nombre: "Morales Luis", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 2.3, lat: -26.544329, lng: -59.357935 },
      { nombre: "Rivarola Erica", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Jueves", institucionId: "sedeChile", kmDesde: 1.9, lat: -26.549676, lng: -59.340585 },
      { nombre: "Robledo Teresa", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Viernes", institucionId: "sedeChile", kmDesde: 2.4, lat: -26.549579, lng: -59.336591 },
      { nombre: "Rotela Alejo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 1.6, lat: -26.529153, lng: -59.345684 },
      { nombre: "Sanchez Andres", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Viernes", institucionId: "sedeChile", kmDesde: 3.4, lat: -26.551097, lng: -59.321861 },
      { nombre: "Silvero Cristian", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 4.0, lat: -26.557319, lng: -59.329471 },
      { nombre: "Soto Faustino", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 3.8, lat: -26.557991, lng: -59.328812 },
      { nombre: "Britez Francisco", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 1.3, lat: -26.540288, lng: -59.352223 },
      { nombre: "Bianchi Sebastian", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 2.3, lat: -26.547318, lng: -59.352405 },
      { nombre: "Britto Federico", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Gral.Jose De San Martin", dias: "Lunes y Jueves", institucionId: "sedeChile", kmDesde: 1.6, lat: -26.531084, lng: -59.338085 },
    ],
  },
  {
    id: "r10", nombre: "Ciervo Laguna", localidad: "Ciervo",
    chofer: "OMAR ORTIZ IDA Y VUELTA", auxiliar: "Sin asignar", vehiculo: "", patente: "AC 617 JG",
    kmIda: 51.6,
    chicos: [
      { nombre: "Saravia Ulises", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Ciervo", dias: "JUEVES", institucionId: "sedeChile", kmDesde: 50.7, lat: -26.582647, lng: -59.634045 },
      { nombre: "Ortiz Jonathan", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Ciervo", dias: "JUEVES", institucionId: "sedeChile", kmDesde: 51.6, lat: -26.580866, lng: -59.624668 },
    ],
  },
  {
    id: "r11", nombre: "Anexo La Leonesa", localidad: "Bermejo, Las Palmas, La Leonesa, General Vedia y otras",
    chofer: "NUÑEZ GASTÓN (Primer turno) - SANCHEZ LUCIO (Segundo turno)", auxiliar: "Sin asignar", vehiculo: "", patente: "OMV 701",
    kmIda: 0,
    chicos: [
      { nombre: "Avalos, Damaris", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Bermejo", dias: "Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Bordon Balcaza, Astor", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Colman Avalos, Uriel", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Cuevas, Gael", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Bermejo", dias: "Viernes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Delpech, Itzel", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "General Vedia", dias: "Martes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Delpech, Naschly", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "General Vedia", dias: "Lunes- Miércoles- Viernes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Dominguez Genes, Ethan", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "General Vedia", dias: "Martes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Fernandez, Gael", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Lunes- Miércoles", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Fleitas, Ramiro", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Miercoles- Viernes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Larrea, Noah", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Lunes- Miércoles", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Lopez Escobar, Mateo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Miercoles- Viernes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Morinigo, Celene", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Lunes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Ojeda, Lautaro", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "General Vedia (Campo)", dias: "Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Ojeda, Samanta", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Martes", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Paez Cabrera, Ramiro", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Lunes- Miércoles", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Rolon, David", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "General Vedia", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Romero, Yisenia", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Torales, Bautista", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Vallejos Belozo, Maximo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "Miércoles", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Villan, Michael", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "Lunes- Miércoles", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Wirz, Lautaro", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "General Vedia", dias: "Martes- Jueves", institucionId: "sedeChile", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
    ],
  },
  {
    id: "r12", nombre: "Las Palmas- La Leonesa", localidad: "La Leonesa, Las Palmas, La Laonesa",
    chofer: "González Gonzalo", auxiliar: "Escalante Sandra", vehiculo: "", patente: "AB 454 LE",
    kmIda: 130.0,
    chicos: [
      { nombre: "Benitez Sergio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 110.0, lat: -27.045097, lng: -58.703563 },
      { nombre: "Candia Santiago", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.032968, lng: -58.698988 },
      { nombre: "Cazal Jesus", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 114.0, lat: -27.061495, lng: -58.675048 },
      { nombre: "Chavez Ramona", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 113.0, lat: -27.059772, lng: -58.678584 },
      { nombre: "Cordoba Anonella", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 110.0, lat: -27.033493, lng: -58.698503 },
      { nombre: "Ganzalez Marcelo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 130.0, lat: -27.049559, lng: -58.67338 },
      { nombre: "Gomez Soto Horacio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 109.0, lat: -27.029325, lng: -58.712471 },
      { nombre: "Gonzalez Cancio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 110.0, lat: -27.047831, lng: -58.705329 },
      { nombre: "Lopez Sebastian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 110.0, lat: -27.037399, lng: -58.697161 },
      { nombre: "Mazeo Naiara Jocelyn", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Laonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 110.0, lat: -27.037696, lng: -58.697491 },
      { nombre: "Montañez Veronica", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.040828, lng: -58.692439 },
      { nombre: "Morinigo Mateo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 113.0, lat: -27.024621, lng: -58.681412 },
      { nombre: "Navarro Margarita", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.031181, lng: -58.703263 },
      { nombre: "Pintos Joselin", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 117.0, lat: -27.018855, lng: -58.718502 },
      { nombre: "Quiroz Andres", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.039468, lng: -58.689252 },
      { nombre: "Romero Emanuel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.032603, lng: -58.699022 },
      { nombre: "Soto Jemena", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Las Palmas", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 121.0, lat: -26.95179, lng: -58.691565 },
      { nombre: "Velezco Saavedra Cristina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Leonesa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 111.0, lat: -27.042216, lng: -58.693109 },
    ],
  },
  {
    id: "r13", nombre: "Elisa Solari", localidad: "Colonia Elisa, Capitán Solari, Elisa",
    chofer: "BENÍTEZ PABLO", auxiliar: "ALVAREZ CARINA (IDA Y VUELTA)", vehiculo: "", patente: "AB 729 WY",
    kmIda: 68.8,
    chicos: [
      { nombre: "Aguirrez Eugenia", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 68.8, lat: -26.926769, lng: -59.51997 },
      { nombre: "Álvarez Yanina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "Lunes a viernes", institucionId: "sedeUruguay", kmDesde: 67.3, lat: -26.931765, lng: -59.518159 },
      { nombre: "Ayala Sergio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "lunes a viernes", institucionId: "sedeUruguay", kmDesde: 65.4, lat: -26.927583, lng: -59.515288 },
      { nombre: "Benítez Ricardo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "lunes a viernes", institucionId: "sedeUruguay", kmDesde: 41.5, lat: -26.74713, lng: -59.614562 },
      { nombre: "Castillo Juan Pablo", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "Lunes y Miercoles", institucionId: "sedeChile", kmDesde: 67.0, lat: -26.936276, lng: -59.512155 },
      { nombre: "Gomez Cristian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "LUNES  Y VIERNES", institucionId: "sedeUruguay", kmDesde: 64.7, lat: -26.932714, lng: -59.523064 },
      { nombre: "Lopez Luis", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 50.3, lat: -26.799995, lng: -59.558182 },
      { nombre: "López Analia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 50.3, lat: -26.799995, lng: -59.558182 },
      { nombre: "López Denis", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Colonia Elisa", dias: "lun-mierc- vier", institucionId: "sedeUruguay", kmDesde: 67.6, lat: -26.93488, lng: -59.515382 },
      { nombre: "Loto Georgina", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "lunes a viernes", institucionId: "sedeUruguay", kmDesde: 50.3, lat: -26.800932, lng: -59.557575 },
      { nombre: "Mendoza Francisca", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "MARTES Y JUEVES", institucionId: "sedeUruguay", kmDesde: 57.8, lat: -26.82699, lng: -59.513936 },
      { nombre: "Núñez Aldana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 50.3, lat: -26.798152, lng: -59.554803 },
      { nombre: "Paz Walter", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Elisa", dias: "Lunes a viernes", institucionId: "sedeUruguay", kmDesde: 67.6, lat: -26.92617, lng: -59.518198 },
      { nombre: "Sanchez Leonor", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "Lunes a viernes", institucionId: "sedeUruguay", kmDesde: 49.9, lat: -26.803504, lng: -59.558237 },
      { nombre: "Silvero Indira", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 50.7, lat: -26.803031, lng: -59.557321 },
      { nombre: "Toledo Aguedo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "lunes Y miercoles", institucionId: "sedeUruguay", kmDesde: 55.1, lat: -26.80705, lng: -59.560375 },
      { nombre: "Zelarayan Cecilia", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Capitán Solari", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 51.5, lat: -26.807647, lng: -59.555471 },
    ],
  },
  {
    id: "r14", nombre: "El Colorado", localidad: "El Colorado, Villa 213",
    chofer: "Bazán, Marcos", auxiliar: "Nuske Ivana (IDA) - Villaba Fanny (VUELTA)", vehiculo: "", patente: "AD 647 MG",
    kmIda: 48.2,
    chicos: [
      { nombre: "Benítez Gabriela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 28.6, lat: -26.312017, lng: -59.358772 },
      { nombre: "Cardozo Nicolas", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "Lunes y miercoles", institucionId: "sedeUruguay", kmDesde: 44.8, lat: -26.191599, lng: -59.361717 },
      { nombre: "Caseres Roman", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 48.2, lat: -26.201405, lng: -59.293793 },
      { nombre: "Fernandez Rene", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 27.1, lat: -26.319143, lng: -59.367088 },
      { nombre: "Figueredo Andrea Soledad", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 29.1, lat: -26.312786, lng: -59.365677 },
      { nombre: "Garcete Segio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 32.4, lat: -26.306578, lng: -59.378727 },
      { nombre: "Hairbar German", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 31.2, lat: -26.299763, lng: -59.367382 },
      { nombre: "Iachini  Matias", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 29.0, lat: -26.305553, lng: -59.36026 },
      { nombre: "Lopez Sergio", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 27.2, lat: -26.319063, lng: -59.368229 },
      { nombre: "Martina Juan Pablo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a viernes", institucionId: "sedeUruguay", kmDesde: 30.5, lat: -26.302763, lng: -59.361858 },
      { nombre: "Martinez Victor", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 46.8, lat: -26.181824, lng: -59.368748 },
      { nombre: "Monzon Javier", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 28.5, lat: -26.314688, lng: -59.35833 },
      { nombre: "Monzon Marianela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 29.2, lat: -26.312717, lng: -59.366859 },
      { nombre: "Ortellado Yanela", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Martes", institucionId: "sedeUruguay", kmDesde: 28.8, lat: -26.312567, lng: -59.36121 },
      { nombre: "Pinter Nelson", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "Lun,mier,vier", institucionId: "sedeUruguay", kmDesde: 29.5, lat: -26.320433, lng: -59.362553 },
      { nombre: "Scheffler Jonathan", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "Martes y Jueves", institucionId: "sedeUruguay", kmDesde: 0, lat: -26.536884, lng: -59.333343 },
      { nombre: "Szachraj Facundo", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 29.9, lat: -26.310429, lng: -59.371395 },
      { nombre: "Velazco Chamorro Diego", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Villa 213", dias: "Lunes a viernes", institucionId: "sedeUruguay", kmDesde: 45.1, lat: -26.190481, lng: -59.364212 },
      { nombre: "Vergara Wlliams", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 26.9, lat: -26.319027, lng: -59.372376 },
      { nombre: "Vicente Eliana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "El Colorado", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 29.4, lat: -26.316128, lng: -59.365292 },
    ],
  },
  {
    id: "r15", nombre: "San Martin (Set)", localidad: "Gral Jose De San Martin",
    chofer: "Sin asignar", auxiliar: "Sin asignar", vehiculo: "", patente: "-",
    kmIda: 7.5,
    chicos: [
      { nombre: "Erick Gomez", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Martes, jueves y viernes", institucionId: "sedeChile", kmDesde: 3.8, lat: -26.519794, lng: -59.364719 },
      { nombre: "Gimenez Franccesca", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Lunes, miercoles y viernes", institucionId: "sedeChile", kmDesde: 1.5, lat: -26.543166, lng: -59.349801 },
      { nombre: "Castillo Ibrahim", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Lunes y jueves", institucionId: "sedeChile", kmDesde: 1.3, lat: -26.535074, lng: -59.350768 },
      { nombre: "Peñalver Nahuel", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Martes y Jueves", institucionId: "sedeChile", kmDesde: 7.5, lat: -26.524242, lng: -59.330967 },
      { nombre: "Rodriguez Abigail", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Lunes, miercoles y viernes", institucionId: "sedeChile", kmDesde: 1.9, lat: -26.547699, lng: -59.333912 },
      { nombre: "Ramirez Aldana", dni: "", prestacion: "SET", obraSocial: "", domicilio: "", localidad: "Gral Jose De San Martin", dias: "Lunes, miercoles y viernes", institucionId: "sedeChile", kmDesde: 1.5, lat: -26.542866, lng: -59.34951 },
    ],
  },
  {
    id: "r16", nombre: "La Escondida - La Verde - Lapachito- Elisa- Unidas", localidad: "La Verde, La Escondida, Lapachito, Elisa",
    chofer: "ROJAS CRISTIAN", auxiliar: "PEDEDILA LUISA (IDA Y VUELTA)", vehiculo: "", patente: "AE 910 ID",
    kmIda: 134.0,
    chicos: [
      { nombre: "Avalos Leandro", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 133.0, lat: -27.122757, lng: -59.37813 },
      { nombre: "Barraza Mirian", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 86.6, lat: -27.103034, lng: -59.448334 },
      { nombre: "Benitez Susana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 134.0, lat: -27.120052, lng: -59.390277 },
      { nombre: "Benitez Victor", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Lapachito", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 128.0, lat: -27.1587, lng: -59.389402 },
      { nombre: "Dominguez Salvador", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "Jueves", institucionId: "sedeChile", kmDesde: 102.0, lat: -27.122046, lng: -59.37832 },
      { nombre: "Fernandez Luciana", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 133.0, lat: -27.12178, lng: -59.381683 },
      { nombre: "Frias María Itati", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 133.0, lat: -27.13235, lng: -59.37736 },
      { nombre: "Lopez Brianna", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "LUNES", institucionId: "sedeChile", kmDesde: 88.2, lat: -27.097797, lng: -59.435287 },
      { nombre: "Ojeda Miguel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "Elisa", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 53.5, lat: -26.825076, lng: -59.54366 },
      { nombre: "Pededila Gabriel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 133.0, lat: -27.131525, lng: -59.374336 },
      { nombre: "Ponce Javier", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 87.4, lat: -27.11174, lng: -59.444581 },
      { nombre: "Reigel Fernando", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 87.4, lat: -27.113641, lng: -59.450277 },
      { nombre: "Reyero Tomas", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "Lunes a Viernes", institucionId: "sedeUruguay", kmDesde: 104.0, lat: -27.121226, lng: -59.390716 },
      { nombre: "Rios Mario", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 86.9, lat: -27.100744, lng: -59.447161 },
      { nombre: "Romero Edrick", dni: "", prestacion: "SAIE", obraSocial: "", domicilio: "", localidad: "La Escondida", dias: "LUNES", institucionId: "sedeChile", kmDesde: 87.1, lat: -27.099304, lng: -59.438205 },
      { nombre: "Sanchez Emmanuel", dni: "", prestacion: "Centro de Día", obraSocial: "", domicilio: "", localidad: "La Verde", dias: "LUNES A VIERNES", institucionId: "sedeUruguay", kmDesde: 133.0, lat: -27.13119, lng: -59.375581 },
    ],
  },
];
const USUARIOS_INICIAL = [
  { id: "u1", nombre: "Maldonado Omar", rol: "Chofer", recorridoId: "r1", usuario: "momar", password: "1234" },
  { id: "u2", nombre: "Zacarias Gustavo", rol: "Auxiliar", recorridoId: "r1", usuario: "zgustavo", password: "1234" },
  { id: "u3", nombre: "Blume Osvaldo", rol: "Chofer", recorridoId: "r2", usuario: "bosvaldo", password: "1234" },
];

/* ---------- pequeños componentes reutilizables ---------- */

function Field({ label, value, onChange, type = "text" }) {
  return (
    <label className="block">
      <span className="text-[11px] uppercase tracking-wide text-[#8B978F]">{label}</span>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(type === "number" ? Number(e.target.value) : e.target.value)}
        className="mt-1 w-full text-[13.5px] rounded-lg border border-[#E2E7E2] px-3 py-2 outline-none focus:border-[#3F6C51]"
      />
    </label>
  );
}

function SelectField({ label, value, onChange, options }) {
  return (
    <label className="block">
      <span className="text-[11px] uppercase tracking-wide text-[#8B978F]">{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)}
        className="mt-1 w-full text-[13.5px] rounded-lg border border-[#E2E7E2] px-3 py-2 outline-none focus:border-[#3F6C51] bg-white">
        {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </label>
  );
}

function Modal({ title, children, onClose, onSave, onDelete }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-black/30 px-0 sm:px-4">
      <div className="bg-white w-full sm:max-w-[440px] sm:rounded-2xl rounded-t-2xl max-h-[88vh] overflow-y-auto">
        <div className="flex items-center justify-between px-5 py-4 border-b border-[#E2E7E2] sticky top-0 bg-white">
          <h3 style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }} className="text-[16px]">{title}</h3>
          <button onClick={onClose}><X size={18} color={MUTED} /></button>
        </div>
        <div className="px-5 py-4 space-y-3">{children}</div>
        <div className="px-5 py-4 flex items-center justify-between gap-2 border-t border-[#E2E7E2]">
          {onDelete ? (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-[13px] text-[#B5533E] px-2 py-2">
              <Trash2 size={14} /> Eliminar
            </button>
          ) : <span />}
          <div className="flex gap-2">
            <button onClick={onClose} className="text-[13px] px-4 py-2 rounded-lg border border-[#E2E7E2]" style={{ color: INK }}>Cancelar</button>
            <button onClick={onSave} className="text-[13px] px-4 py-2 rounded-lg text-white" style={{ background: GREEN }}>Guardar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function Stat({ icon: Icon, label, value, unit, accent }) {
  return (
    <div className="flex-1 min-w-[140px] rounded-xl bg-white border border-[#E2E7E2] px-4 py-3">
      <div className="flex items-center gap-1.5 text-[#5B6B63]">
        <Icon size={14} strokeWidth={2} />
        <span className="text-[11px] tracking-wide uppercase">{label}</span>
      </div>
      <div className="mt-1.5 flex items-baseline gap-1">
        <span className="text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: accent || INK, fontWeight: 500 }}>{value}</span>
        <span className="text-xs text-[#8B978F]">{unit}</span>
      </div>
    </div>
  );
}

/* ---------- diagrama y mapa (ahora con múltiples instituciones) ---------- */

function RouteDiagram({ recorrido, instituciones }) {
  const institucion = (id) => instituciones.find((i) => i.id === id);

  // agrupa las entregas por institución, en el orden en que aparece cada una por primera vez
  const grupos = [];
  recorrido.chicos.forEach((c) => {
    let g = grupos.find((g) => g.institucionId === c.institucionId);
    if (!g) { g = { institucionId: c.institucionId, chicos: [] }; grupos.push(g); }
    g.chicos.push(c);
  });

  return (
    <div>
      <div className="text-[10.5px] uppercase tracking-wide mb-2" style={{ color: MUTED }}>Recogida</div>
      <div className="relative pl-1">
        <div className="absolute left-[9px] top-2 bottom-2 w-px bg-[#D7DED7]" />
        {recorrido.chicos.map((c, i) => (
          <div key={i} className="relative flex items-start gap-4 pb-5 last:pb-0">
            <div className="relative z-10 mt-0.5 shrink-0 rounded-full flex items-center justify-center"
              style={{ width: 19, height: 19, background: "#FFFFFF", border: `2px solid ${GREEN}` }} />
            <div className="flex-1 flex items-center justify-between gap-3">
              <div>
                <div className="text-[13.5px] leading-tight" style={{ fontWeight: 500, color: INK }}>{c.nombre}</div>
                <div className="text-[11.5px] mt-0.5" style={{ color: MUTED }}>{c.domicilio}</div>
              </div>
              <span className="text-[11.5px] shrink-0 rounded-full px-2 py-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace", color: OCHRE, background: "#FBF1E6" }}>
                +{c.kmDesde.toFixed(1)} km
              </span>
            </div>
          </div>
        ))}
      </div>

      <div className="text-[10.5px] uppercase tracking-wide mt-5 mb-2" style={{ color: MUTED }}>Entrega en instituciones</div>
      <div className="space-y-3">
        {grupos.map((g, i) => {
          const inst = institucion(g.institucionId);
          return (
            <div key={i} className="rounded-lg px-3 py-2.5" style={{ background: "#EEF3EC" }}>
              <div className="flex items-center gap-1.5">
                <Building2 size={13} color={GREEN} />
                <span className="text-[13px]" style={{ fontWeight: 500, color: "#2C4A34" }}>{inst ? inst.nombre : "Institución sin asignar"}</span>
              </div>
              <div className="text-[11.5px] mt-1" style={{ color: "#4C6152" }}>
                {g.chicos.map((c) => c.prestacion ? `${c.nombre} (${c.prestacion})` : c.nombre).join(" · ")}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function projectLatLng(points, viewW = 100, viewH = 75, pad = 10) {
  const lats = points.map((p) => p.lat);
  const lngs = points.map((p) => p.lng);
  const minLat = Math.min(...lats), maxLat = Math.max(...lats);
  const minLng = Math.min(...lngs), maxLng = Math.max(...lngs);
  const avgLat = (minLat + maxLat) / 2;
  const cosLat = Math.cos((avgLat * Math.PI) / 180);
  const lngSpan = Math.max((maxLng - minLng) * cosLat, 0.0005);
  const latSpan = Math.max(maxLat - minLat, 0.0005);
  const usableW = viewW - 2 * pad, usableH = viewH - 2 * pad;
  const scale = Math.min(usableW / lngSpan, usableH / latSpan);
  const offsetX = pad + (usableW - lngSpan * scale) / 2;
  const offsetY = pad + (usableH - latSpan * scale) / 2;
  return points.map((p) => ({
    x: offsetX + (p.lng - minLng) * cosLat * scale,
    y: offsetY + (maxLat - p.lat) * scale,
  }));
}

function RouteMap({ recorrido, instituciones }) {
  const mapDivRef = React.useRef(null);
  const [status, setStatus] = useState("loading"); // loading | ready | error
  const [errorMsg, setErrorMsg] = useState("");
  const [routeWarning, setRouteWarning] = useState("");

  const institucionesUsadas = [];
  recorrido.chicos.forEach((c) => {
    if (!institucionesUsadas.find((i) => i.id === c.institucionId)) {
      const inst = instituciones.find((i) => i.id === c.institucionId);
      if (inst) institucionesUsadas.push(inst);
    }
  });

  useEffect(() => {
    let cancelled = false;

    loadGoogleMaps()
      .then((maps) => {
        if (cancelled || !mapDivRef.current) return;
        if (recorrido.chicos.length === 0) { setStatus("ready"); return; }

        const map = new maps.Map(mapDivRef.current, {
          zoom: 12,
          center: { lat: recorrido.chicos[0].lat, lng: recorrido.chicos[0].lng },
          disableDefaultUI: false,
          mapTypeControl: false,
          streetViewControl: false,
        });

        const bounds = new maps.LatLngBounds();
        const markers = [];

        // marcadores de domicilios — se numeran después, según el orden optimizado
        recorrido.chicos.forEach((c) => {
          const pos = { lat: c.lat, lng: c.lng };
          bounds.extend(pos);
          const marker = new maps.Marker({
            position: pos, map,
            icon: { path: maps.SymbolPath.CIRCLE, scale: 11, fillColor: "#3F6C51", fillOpacity: 1, strokeColor: "#FFFFFF", strokeWeight: 2 },
            title: c.nombre,
          });
          markers.push(marker);
        });

        const setMarkerNumber = (marker, n) => {
          marker.setLabel({ text: String(n), color: "#FFFFFF", fontSize: "11px", fontWeight: "600" });
        };
        // numeración por defecto (orden de la planilla) hasta que llegue la optimizada
        markers.forEach((m, i) => setMarkerNumber(m, i + 1));

        // marcadores de instituciones (sedes) involucradas
        institucionesUsadas.forEach((inst) => {
          const pos = { lat: inst.lat, lng: inst.lng };
          bounds.extend(pos);
          new maps.Marker({
            position: pos, map,
            icon: { path: maps.SymbolPath.BACKWARD_CLOSED_ARROW, scale: 5, rotation: 0, fillColor: "#C97B3D", fillOpacity: 1, strokeColor: "#FFFFFF", strokeWeight: 1.5 },
            title: inst.nombre,
          });
        });

        map.fitBounds(bounds);

        if (recorrido.chicos.length < 2 && institucionesUsadas.length === 0) {
          setStatus("ready");
          return;
        }

        const domicilios = recorrido.chicos.map((c) => ({ lat: c.lat, lng: c.lng }));
        const destinoInst = institucionesUsadas[institucionesUsadas.length - 1];
        const destination = destinoInst ? { lat: destinoInst.lat, lng: destinoInst.lng } : domicilios[domicilios.length - 1];

        // el origen es el primer domicilio; los "waypoints" son el resto de los domicilios
        // (todos, si el destino es una sede; todos menos el último, si el destino es el último domicilio)
        const origin = domicilios[0];
        const waypointOriginalIndices = destinoInst
          ? domicilios.map((_, i) => i).slice(1)
          : domicilios.map((_, i) => i).slice(1, -1);
        const waypoints = waypointOriginalIndices.map((i) => ({ location: domicilios[i], stopover: true }));

        const directionsService = new maps.DirectionsService();
        const directionsRenderer = new maps.DirectionsRenderer({
          map, suppressMarkers: true,
          polylineOptions: { strokeColor: "#3F6C51", strokeWeight: 4, strokeOpacity: 0.85 },
        });

        directionsService.route(
          { origin, destination, waypoints, travelMode: maps.TravelMode.DRIVING, optimizeWaypoints: true },
          (result, statusResult) => {
            if (cancelled) return;
            if (statusResult === "OK") {
              directionsRenderer.setDirections(result);
              // renumerar los marcadores según el orden optimizado que devolvió Google
              const order = result.routes[0].waypoint_order || waypointOriginalIndices.map((_, i) => i);
              const sequence = [0, ...order.map((i) => waypointOriginalIndices[i])];
              if (!destinoInst) sequence.push(domicilios.length - 1);
              sequence.forEach((origIdx, pos) => setMarkerNumber(markers[origIdx], pos + 1));
              setRouteWarning("");
            } else {
              console.error("Directions API status:", statusResult, result);
              setRouteWarning(
                statusResult === "ZERO_RESULTS"
                  ? "No se encontró un camino por calles entre esos puntos."
                  : statusResult === "REQUEST_DENIED"
                  ? "La API de Direcciones no está habilitada o la clave no tiene permiso (revisar en Google Cloud)."
                  : statusResult === "OVER_QUERY_LIMIT"
                  ? "Se alcanzó el límite de consultas de la API por ahora."
                  : `No se pudo calcular la ruta (${statusResult}). Se muestran los puntos sin trazado.`
              );
              // conexión de respaldo: línea recta entre los puntos, en el orden de la planilla
              new maps.Polyline({
                map,
                path: [...domicilios, destination],
                strokeColor: "#B9CBBA",
                strokeOpacity: 0.9,
                strokeWeight: 3,
                icons: [{ icon: { path: "M 0,-1 0,1", strokeOpacity: 1, scale: 3 }, offset: "0", repeat: "14px" }],
              });
            }
            setStatus("ready");
          }
        );
      })
      .catch((err) => {
        if (!cancelled) { setStatus("error"); setErrorMsg(err.message); }
      });

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [recorrido.id]);

  if (status === "error") {
    return (
      <div className="w-full rounded-lg flex items-center justify-center text-center px-4" style={{ background: "#FBEEEC", aspectRatio: "4 / 3" }}>
        <p className="text-[12.5px]" style={{ color: "#B5533E" }}>No se pudo cargar el mapa: {errorMsg}</p>
      </div>
    );
  }

  return (
    <div className="relative w-full rounded-lg overflow-hidden" style={{ aspectRatio: "4 / 3" }}>
      {status === "loading" && (
        <div className="absolute inset-0 flex items-center justify-center" style={{ background: "#EAF0EA" }}>
          <p className="text-[12.5px]" style={{ color: MUTED }}>Cargando mapa...</p>
        </div>
      )}
      <div ref={mapDivRef} className="w-full h-full" />
      {routeWarning && (
        <div className="absolute left-2 right-2 bottom-2 text-[11px] rounded-md px-2.5 py-1.5" style={{ background: "#FBEEEC", color: "#B5533E" }}>
          {routeWarning}
        </div>
      )}
    </div>
  );
}

const emptyRecorrido = () => ({ nombre: "", localidad: "", chofer: "", auxiliar: "", vehiculo: "", patente: "", kmIda: 0, chicos: [] });
const emptyChico = (instituciones, prestaciones, localidades, obrasSociales) => {
  const pres = prestaciones && prestaciones[0];
  const institucionId = pres ? pres.institucionId : instituciones[0]?.id || "";
  const inst = instituciones.find((i) => i.id === institucionId) || instituciones[0];
  return {
    nombre: "", dni: "", prestacion: pres ? pres.nombre : "",
    obraSocial: obrasSociales && obrasSociales[0] ? obrasSociales[0].nombre : "",
    domicilio: "", localidad: localidades && localidades[0] ? localidades[0].nombre : "",
    dias: "", institucionId, kmDesde: 0, lat: inst?.lat, lng: inst?.lng, coordsManual: "",
  };
};
const emptyUsuario = (recorridos) => ({ nombre: "", rol: "Chofer", recorridoId: recorridos[0]?.id || "", usuario: "", password: "" });
const emptyInstitucion = () => ({ nombre: "", direccion: "", lat: -33.08387, lng: -68.47312 });
const emptyPrestacion = (instituciones) => ({ nombre: "", institucionId: instituciones[0]?.id || "" });
const emptyLocalidad = () => ({ nombre: "", provincia: "" });
const emptyObraSocial = () => ({ nombre: "" });

/* ---------- pantalla de inicio ---------- */

function LoginDecoration() {
  return (
    <svg viewBox="0 0 500 700" className="absolute inset-0 w-full h-full" preserveAspectRatio="xMidYMid slice">
      <defs>
        <linearGradient id="loginBg" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#2C4A34" />
          <stop offset="100%" stopColor="#3F6C51" />
        </linearGradient>
      </defs>
      <rect width="500" height="700" fill="url(#loginBg)" />
      <circle cx="60" cy="620" r="180" fill="#FFFFFF" opacity="0.06" />
      <circle cx="460" cy="80" r="140" fill="#F7C948" opacity="0.10" />
      <circle cx="420" cy="520" r="90" fill="#FFFFFF" opacity="0.05" />
      {/* ruta punteada con paradas, ecoando el mapa de recorridos */}
      <path d="M 70 560 Q 160 480 150 380 T 260 220 T 250 90" fill="none" stroke="#FFFFFF" strokeOpacity="0.35" strokeWidth="3" strokeDasharray="2 14" strokeLinecap="round" />
      <circle cx="70" cy="560" r="9" fill="#F7C948" />
      <circle cx="150" cy="380" r="6" fill="#FFFFFF" fillOpacity="0.9" />
      <circle cx="260" cy="220" r="6" fill="#FFFFFF" fillOpacity="0.9" />
      <circle cx="250" cy="90" r="10" fill="#FFFFFF" />
      <circle cx="250" cy="90" r="10" fill="none" stroke="#FFFFFF" strokeWidth="2" opacity="0.6">
        <animate attributeName="r" values="10;22;10" dur="2.8s" repeatCount="indefinite" />
        <animate attributeName="opacity" values="0.6;0;0.6" dur="2.8s" repeatCount="indefinite" />
      </circle>
    </svg>
  );
}

function Login({ onEnter, usuarios }) {
  const [role, setRole] = useState("admin");
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  function handleSubmit(e) {
    e.preventDefault();
    setError("");
    if (role === "admin") {
      onEnter("admin", null);
      return;
    }
    const match = usuarios.find(
      (u) => u.usuario.trim().toLowerCase() === usuario.trim().toLowerCase() && (u.password || "") === password
    );
    if (match) {
      onEnter("usuario", match);
    } else {
      setError("Usuario o contraseña incorrectos.");
    }
  }

  return (
    <div className="min-h-screen w-full flex" style={{ background: "#F3F5F1" }}>
      <style>{FONT_IMPORT}</style>

      {/* Panel decorativo, oculto en mobile */}
      <div className="hidden md:flex md:w-[46%] relative overflow-hidden items-center justify-center">
        <LoginDecoration />
        <div className="relative z-10 text-center px-10">
          <img src={LOGO_URL} alt="Despertares" className="h-16 mx-auto object-contain" style={{ filter: "brightness(0) invert(1)" }} />
          <h2 className="mt-6 text-[22px] text-white" style={{ fontFamily: "Fraunces", fontWeight: 500 }}>
            Cada recorrido, bajo control
          </h2>
          <p className="mt-2 text-[13px] text-white" style={{ opacity: 0.85 }}>
            Transporte, recorridos e instituciones en un solo lugar
          </p>
        </div>
      </div>

      {/* Formulario */}
      <div className="flex-1 flex items-center justify-center px-6 py-10">
        <div className="w-full max-w-[380px]">
          <img src={LOGO_URL} alt="Despertares" className="h-14 mx-auto object-contain md:hidden mb-6" />

          <h1 className="text-[24px] text-center md:text-left" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>
            Iniciar sesión
          </h1>
          <p className="mt-1.5 text-[13px] text-center md:text-left" style={{ color: MUTED }}>
            Ingresá con tu usuario para continuar
          </p>

          <div className="mt-6 flex rounded-full p-1" style={{ background: "#EEF3EC" }}>
            <button onClick={() => { setRole("admin"); setError(""); }} className="flex-1 flex items-center justify-center gap-1.5 text-[12.5px] rounded-full py-2"
              style={{ background: role === "admin" ? "#FFFFFF" : "transparent", color: role === "admin" ? GREEN : "#5B6B63", fontWeight: role === "admin" ? 600 : 400, boxShadow: role === "admin" ? "0 1px 2px rgba(0,0,0,0.06)" : "none" }}>
              <Settings size={13} /> Administrador
            </button>
            <button onClick={() => { setRole("usuario"); setError(""); }} className="flex-1 flex items-center justify-center gap-1.5 text-[12.5px] rounded-full py-2"
              style={{ background: role === "usuario" ? "#FFFFFF" : "transparent", color: role === "usuario" ? GREEN : "#5B6B63", fontWeight: role === "usuario" ? 600 : 400, boxShadow: role === "usuario" ? "0 1px 2px rgba(0,0,0,0.06)" : "none" }}>
              <Bus size={13} /> Chofer / Auxiliar
            </button>
          </div>

          <form onSubmit={handleSubmit} className="mt-6 space-y-3.5">
            <label className="block">
              <span className="text-[11px] uppercase tracking-wide" style={{ color: MUTED }}>Usuario</span>
              <div className="mt-1 flex items-center gap-2 rounded-lg border border-[#E2E7E2] px-3 py-2.5 focus-within:border-[#3F6C51]">
                <User size={15} color={MUTED} />
                <input value={usuario} onChange={(e) => setUsuario(e.target.value)} placeholder={role === "admin" ? "admin" : "ej. momar"}
                  className="w-full text-[13.5px] outline-none" style={{ color: INK }} />
              </div>
            </label>
            <label className="block">
              <span className="text-[11px] uppercase tracking-wide" style={{ color: MUTED }}>Contraseña</span>
              <div className="mt-1 flex items-center gap-2 rounded-lg border border-[#E2E7E2] px-3 py-2.5 focus-within:border-[#3F6C51]">
                <Lock size={15} color={MUTED} />
                <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••"
                  className="w-full text-[13.5px] outline-none" style={{ color: INK }} />
              </div>
            </label>

            {error && <p className="text-[12px]" style={{ color: "#B5533E" }}>{error}</p>}

            <div className="flex items-center justify-between pt-1">
              <label className="flex items-center gap-2 text-[12px]" style={{ color: MUTED }}>
                <input type="checkbox" className="accent-[#3F6C51] w-3.5 h-3.5" /> Recordarme
              </label>
              <button type="button" className="text-[12px]" style={{ color: GREEN }}>¿Olvidaste tu contraseña?</button>
            </div>

            <button type="submit" className="w-full rounded-full py-3 text-[13.5px] text-white mt-2" style={{ background: GREEN, fontWeight: 500 }}>
              Ingresar
            </button>
          </form>

          <p className="mt-6 text-[11px] text-center md:text-left leading-relaxed" style={{ color: MUTED }}>
            {role === "admin"
              ? "Acceso completo: recorridos, usuarios, datos generales, importar y exportar."
              : "Vas a ver únicamente el recorrido que tenés asignado, con su mapa, para guiarte aunque estés reemplazando a otro chofer."}
          </p>
        </div>
      </div>
    </div>
  );
}

function Landing({ onEnter }) {
  return (
    <div className="min-h-screen w-full flex items-center justify-center px-6" style={{ background: "#F3F5F1" }}>
      <style>{FONT_IMPORT}</style>
      <div className="w-full max-w-[420px] text-center">
        <img src={LOGO_URL} alt="Despertares" className="h-16 mx-auto object-contain" />
        <h1 className="mt-5 text-[26px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>
          Transporte de concurrentes
        </h1>
        <p className="mt-2 text-[13.5px]" style={{ color: MUTED }}>
          Elegí el turno para ver los recorridos correspondientes
        </p>
        <div className="mt-8 grid grid-cols-1 gap-3">
          <button onClick={() => onEnter("mañana")}
            className="flex items-center gap-4 rounded-2xl bg-white border border-[#E2E7E2] px-5 py-5 text-left hover:border-[#C9DBC9] transition-colors">
            <div className="w-11 h-11 rounded-full flex items-center justify-center shrink-0" style={{ background: "#FBF1E6" }}>
              <Sun size={20} color={OCHRE} />
            </div>
            <div>
              <div className="text-[15px]" style={{ fontWeight: 500, color: INK }}>Turno mañana</div>
              <div className="text-[12px]" style={{ color: MUTED }}>Ingresar a los recorridos del turno mañana</div>
            </div>
          </button>
          <button onClick={() => onEnter("tarde")}
            className="flex items-center gap-4 rounded-2xl bg-white border border-[#E2E7E2] px-5 py-5 text-left hover:border-[#C9DBC9] transition-colors">
            <div className="w-11 h-11 rounded-full flex items-center justify-center shrink-0" style={{ background: "#EEF3EC" }}>
              <Moon size={20} color={GREEN} />
            </div>
            <div>
              <div className="text-[15px]" style={{ fontWeight: 500, color: INK }}>Turno tarde</div>
              <div className="text-[12px]" style={{ color: MUTED }}>Ingresar a los recorridos del turno tarde</div>
            </div>
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- app principal ---------- */

/* ---------- importador de planillas Excel (mismo formato que TRANSPORTE / TURNO / chicos) ---------- */

function normText(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/\u00a0/g, " ").trim();
}
function titleCase(s) {
  return s.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
}
function normInstitucionNombre(text) {
  const t = text.toUpperCase().replace(/\./g, "").trim();
  if (t.includes("SAIE")) return "SAIE";
  if (t.includes("SET")) return "SET";
  if (t.includes("CENTRO")) return "Centro de Día";
  return text.trim() ? titleCase(text.trim()) : "Sin especificar";
}
// Tolera guiones "raros" (en dash, em dash, signo menos matemático) y espacios no separables
// que a veces quedan al pegar coordenadas desde Google Maps u otras fuentes.
function parseCoordsText(text) {
  const cleaned = String(text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[\u2010-\u2015\u2212]/g, "-")
    .trim();
  const m = /^(-?\d+[.,]?\d*)\s*,\s*(-?\d+[.,]?\d*)$/.exec(cleaned);
  if (!m) return null;
  const lat = parseFloat(m[1].replace(",", "."));
  const lng = parseFloat(m[2].replace(",", "."));
  if (isNaN(lat) || isNaN(lng)) return null;
  return { lat, lng };
}
const IMPORT_MARKERS = ["Datos de veh", "Datos chofer", "Datos auxiliar", "Datos Auxiliar", "AUXILIAR:", "TRANSPORTE:", "TURNO:", "Horario de"];
function isImportMarker(text) { return IMPORT_MARKERS.some((m) => text.startsWith(m)); }

function parseTransporteSheet(rows) {
  const recorridos = [];
  const n = rows.length;
  let i = 0;
  while (i < n) {
    const row = rows[i] || [];
    const c2 = normText(row[2]);
    if (c2 === "TRANSPORTE:") {
      let nombre = normText(row[3]) || normText(row[4]);
      i++;
      if (i < n && normText((rows[i] || [])[2]).toUpperCase().startsWith("TURNO:")) i++;
      if (i < n && normText((rows[i] || [])[2]).includes("Horario de salida")) i++;
      if (i < n && normText((rows[i] || [])[2]).includes("Horario de llegada")) i++;
      while (i < n) {
        const c2c = normText((rows[i] || [])[2]);
        if (c2c === "TRANSPORTE:") break;
        if (c2c === "" || c2c.toUpperCase() === "N°" || c2c.toUpperCase() === "NOMBRE Y APELLIDO" || isImportMarker(c2c)) { i++; continue; }
        break;
      }
      const chicos = [];
      while (i < n) {
        const r = rows[i] || [];
        const c = normText(r[2]);
        const d = normText(r[3]);
        if (c === "" || isImportMarker(c)) break;
        if (d === "") { i++; continue; }
        const dias = normText(r[4]);
        const institucionTxt = normText(r[5]);
        // Normalmente las coordenadas están en la columna G (índice 6), pero algunos bloques
        // tienen una columna extra (ej. "OBSERVACIÓN") antes, corriendo todo un lugar. Probamos
        // primero la posición habitual y, si no matchea, buscamos en las columnas vecinas.
        let coordsIdx = null;
        for (const ci of [6, 7, 8, 5, 9]) {
          if (parseCoordsText(normText(r[ci]))) { coordsIdx = ci; break; }
        }
        const coords = normText(r[coordsIdx !== null ? coordsIdx : 6]);
        const kmRaw = r[coordsIdx !== null ? coordsIdx + 1 : 7];
        const parsedCoords = parseCoordsText(coords);
        const lat = parsedCoords ? parsedCoords.lat : null;
        const lng = parsedCoords ? parsedCoords.lng : null;
        const km = typeof kmRaw === "number" ? kmRaw : parseFloat(kmRaw);
        chicos.push({ nombre: titleCase(c), localidad: titleCase(d), dias, institucion: normInstitucionNombre(institucionTxt), lat, lng, km: isNaN(km) ? null : km });
        i++;
      }
      let vehiculo = "", chofer = "", auxiliar = "";
      const stop = i + 10;
      while (i < n && i < stop) {
        const c2b = normText((rows[i] || [])[2]);
        if (c2b.startsWith("Datos de veh")) vehiculo = c2b.includes(":") ? c2b.split(":").slice(1).join(":").trim() : "";
        else if (c2b.startsWith("Datos chofer")) chofer = c2b.includes(":") ? c2b.split(":").slice(1).join(":").trim() : "";
        else if (c2b.startsWith("Datos auxiliar") || c2b.startsWith("Datos Auxiliar") || c2b.startsWith("AUXILIAR:")) auxiliar = c2b.includes(":") ? c2b.split(":").slice(1).join(":").trim() : "";
        else if (c2b === "TRANSPORTE:") break;
        i++;
      }
      if (chicos.length) recorridos.push({ nombre: titleCase(nombre), chofer, auxiliar, patente: vehiculo, chicos });
    } else {
      i++;
    }
  }
  return recorridos;
}

function buildImportResult(workbook, existingPrestaciones, fallbackLat, fallbackLng) {
  let mainSheetName = workbook.SheetNames[0];
  let maxRows = 0;
  workbook.SheetNames.forEach((name) => {
    const ws = workbook.Sheets[name];
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    const rowCount = range.e.r - range.s.r + 1;
    if (rowCount > maxRows) { maxRows = rowCount; mainSheetName = name; }
  });
  const mainRows = XLSX.utils.sheet_to_json(workbook.Sheets[mainSheetName], { header: 1, raw: true, defval: null });
  const parsedRecorridos = parseTransporteSheet(mainRows);

  let lat0 = null, lng0 = null;
  workbook.SheetNames.forEach((name) => {
    if (name === mainSheetName) return;
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: null });
    rows.forEach((r) => {
      const a = normText((r || [])[0]).toUpperCase();
      if (a.includes("INSTITUCION")) {
        const parsed = parseCoordsText(normText(r[1]));
        if (parsed) { lat0 = parsed.lat; lng0 = parsed.lng; }
      }
    });
  });
  if (lat0 === null) { lat0 = fallbackLat; lng0 = fallbackLng; }

  // Esa columna es la PRESTACIÓN, no la sede. Matcheamos contra las prestaciones ya configuradas
  // (por nombre); las que no existan quedan sin sede asignada para revisar en Configuración.
  const nombresPres = [];
  parsedRecorridos.forEach((r) => r.chicos.forEach((c) => { if (!nombresPres.includes(c.institucion)) nombresPres.push(c.institucion); }));
  const prestacionesNuevas = [];
  const presInstMap = {};
  nombresPres.forEach((nombre) => {
    const existente = existingPrestaciones.find((p) => p.nombre.toLowerCase() === nombre.toLowerCase());
    if (existente) {
      presInstMap[nombre] = existente.institucionId;
    } else {
      presInstMap[nombre] = null;
      prestacionesNuevas.push({ id: "imp_p_" + nombre.replace(/\s+/g, "_"), nombre, institucionId: null });
    }
  });

  const outRecorridos = parsedRecorridos.map((r, idx) => {
    const localidades = [];
    r.chicos.forEach((c) => { if (!localidades.includes(c.localidad)) localidades.push(c.localidad); });
    const kms = r.chicos.map((c) => c.km).filter((k) => typeof k === "number");
    const kmIda = kms.length ? Math.max(...kms) : 0;
    return {
      id: "imp_r" + (idx + 1),
      nombre: r.nombre,
      localidad: localidades.slice(0, 4).join(", ") + (localidades.length > 4 ? " y otras" : ""),
      chofer: r.chofer || "Sin asignar",
      auxiliar: r.auxiliar || "Sin asignar",
      vehiculo: "",
      patente: r.patente || "-",
      kmIda: Math.round(kmIda * 10) / 10,
      chicos: r.chicos.map((c) => ({
        nombre: c.nombre, dni: "", prestacion: c.institucion, obraSocial: "", domicilio: "", dias: c.dias,
        institucionId: presInstMap[c.institucion], kmDesde: c.km ? Math.round(c.km * 10) / 10 : 0,
        lat: c.lat ?? lat0, lng: c.lng ?? lng0,
      })),
    };
  });

  const sinCoordsPropias = outRecorridos.reduce(
    (acc, r) => acc + r.chicos.filter((c) => c.lat === lat0 && c.lng === lng0).length,
    0
  );

  return { prestacionesNuevas, recorridos: outRecorridos, sinCoordsPropias };
}

export default function App() {
  const [screen, setScreen] = useState("login");
  const [role, setRole] = useState("admin");
  const [loggedInUsuario, setLoggedInUsuario] = useState(null);
  const [turno, setTurno] = useState(null);
  const [section, setSection] = useState("recorridos");

  const [recorridos, recorridosReady] = useSeededCollection("recorridos", RECORRIDOS_INICIAL);
  const [usuarios] = useSeededCollection("usuarios", USUARIOS_INICIAL);
  const [instituciones] = useSeededCollection("instituciones", INSTITUCIONES_INICIAL);
  const [prestaciones] = useSeededCollection("prestaciones", PRESTACIONES_INICIAL);
  const [localidades] = useSeededCollection("localidades", LOCALIDADES_INICIAL);
  const [obrasSociales] = useSeededCollection("obrasSociales", OBRAS_SOCIALES_INICIAL);
  const [selectedId, setSelectedId] = useState(null);
  useEffect(() => {
    if (!selectedId && recorridos.length > 0) setSelectedId(recorridos[0].id);
  }, [recorridos, selectedId]);
  const [diasPorMes, setDiasPorMes] = useState(20);
  const [query, setQuery] = useState("");
  const [view, setView] = useState("mapa");
  const [expandedChico, setExpandedChico] = useState(null);
  const [recorridosPanelOpen, setRecorridosPanelOpen] = useState(true);
  const [optimizando, setOptimizando] = useState(false);
  const [optimizarMsg, setOptimizarMsg] = useState("");
  const [modal, setModal] = useState(null);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [isDesktop, setIsDesktop] = useState(true);
  const [datosTab, setDatosTab] = useState("instituciones");

  useEffect(() => {
    const mq = window.matchMedia("(min-width: 768px)");
    const update = () => setIsDesktop(mq.matches);
    update();
    mq.addEventListener ? mq.addEventListener("change", update) : mq.addListener(update);
    return () => (mq.removeEventListener ? mq.removeEventListener("change", update) : mq.removeListener(update));
  }, []);
  const [exportRecorridoIds, setExportRecorridoIds] = useState(RECORRIDOS_INICIAL.map((r) => r.id));
  const [exportInstitucionIds, setExportInstitucionIds] = useState(INSTITUCIONES_INICIAL.map((i) => i.id));
  const [exportIncluirTransporte, setExportIncluirTransporte] = useState(true);
  const [importPreview, setImportPreview] = useState(null);
  const [importError, setImportError] = useState("");

  function handleImportFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImportError("");
    setImportPreview(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const fallbackLat = instituciones[0]?.lat ?? 0;
        const fallbackLng = instituciones[0]?.lng ?? 0;
        const result = buildImportResult(wb, prestaciones, fallbackLat, fallbackLng);
        if (result.recorridos.length === 0) {
          setImportError("No se encontraron recorridos con el formato esperado (bloques que empiecen con \"TRANSPORTE:\").");
        } else {
          setImportPreview(result);
        }
      } catch (err) {
        setImportError("No se pudo leer el archivo. Verificá que sea un Excel (.xlsx) con el formato esperado.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  async function confirmImport() {
    if (!importPreview) return;
    if (importPreview.prestacionesNuevas.length > 0) {
      await Promise.all(importPreview.prestacionesNuevas.map(({ id, ...rest }) => saveItem("prestaciones", id, rest)));
    }
    await replaceCollection("recorridos", recorridos.map((r) => r.id), importPreview.recorridos);
    setSelectedId(importPreview.recorridos[0]?.id);
    setImportPreview(null);
    setSection("recorridos");
  }

  const recorrido = recorridos.find((r) => r.id === selectedId) || recorridos[0];
  const kmDiario = useMemo(() => (recorrido ? recorrido.kmIda * 2 : 0), [recorrido]);
  const kmMensual = useMemo(() => kmDiario * diasPorMes, [kmDiario, diasPorMes]);
  const filtered = recorridos.filter((r) => (r.nombre + r.localidad).toLowerCase().includes(query.toLowerCase()));

  function openNuevoRecorrido() { setModal({ type: "recorrido", mode: "new", draft: emptyRecorrido() }); }
  function openEditarRecorrido() { setModal({ type: "recorrido", mode: "edit", draft: { ...recorrido } }); }
  function openNuevoConcurrente() { setModal({ type: "concurrente", mode: "new", draft: emptyChico(instituciones, prestaciones, localidades, obrasSociales) }); }

  async function optimizarOrden() {
    if (!recorrido || recorrido.chicos.length < 3) {
      setOptimizarMsg("Hacen falta al menos 3 concurrentes para optimizar el orden.");
      return;
    }
    setOptimizando(true);
    setOptimizarMsg("");
    try {
      const maps = await loadGoogleMaps();
      const institucionesUsadas = [];
      recorrido.chicos.forEach((c) => {
        if (!institucionesUsadas.find((i) => i.id === c.institucionId)) {
          const inst = instituciones.find((i) => i.id === c.institucionId);
          if (inst) institucionesUsadas.push(inst);
        }
      });
      const domicilios = recorrido.chicos.map((c) => ({ lat: c.lat, lng: c.lng }));
      const destinoInst = institucionesUsadas[institucionesUsadas.length - 1];
      const destination = destinoInst ? { lat: destinoInst.lat, lng: destinoInst.lng } : domicilios[domicilios.length - 1];
      const origin = domicilios[0];
      const waypointOriginalIndices = destinoInst
        ? domicilios.map((_, i) => i).slice(1)
        : domicilios.map((_, i) => i).slice(1, -1);
      const waypoints = waypointOriginalIndices.map((i) => ({ location: domicilios[i], stopover: true }));

      const directionsService = new maps.DirectionsService();
      directionsService.route(
        { origin, destination, waypoints, travelMode: maps.TravelMode.DRIVING, optimizeWaypoints: true },
        async (result, statusResult) => {
          if (statusResult === "OK") {
            const order = result.routes[0].waypoint_order || waypointOriginalIndices.map((_, i) => i);
            const sequence = [0, ...order.map((i) => waypointOriginalIndices[i])];
            if (!destinoInst) sequence.push(domicilios.length - 1);
            const chicosOrdenados = sequence.map((origIdx) => recorrido.chicos[origIdx]);
            const { id: rId, ...rRest } = recorrido;
            await saveItem("recorridos", recorrido.id, { ...rRest, chicos: chicosOrdenados });
            setOptimizarMsg("Listo — el orden se actualizó según la ruta más corta por calles.");
          } else {
            setOptimizarMsg(
              statusResult === "ZERO_RESULTS"
                ? "No se encontró un camino por calles entre esos puntos."
                : `No se pudo optimizar el orden (${statusResult}).`
            );
          }
          setOptimizando(false);
        }
      );
    } catch (err) {
      setOptimizarMsg("No se pudo optimizar: " + err.message);
      setOptimizando(false);
    }
  }
  function openEditarConcurrente(i) { setModal({ type: "concurrente", mode: "edit", index: i, draft: { ...recorrido.chicos[i] } }); }
  function openNuevoUsuario() { setModal({ type: "usuario", mode: "new", draft: emptyUsuario(recorridos) }); }
  function openEditarUsuario(u) { setModal({ type: "usuario", mode: "edit", id: u.id, draft: { ...u } }); }
  function openNuevaInstitucion() { setModal({ type: "institucion", mode: "new", draft: emptyInstitucion() }); }
  function openEditarInstitucion(i) { setModal({ type: "institucion", mode: "edit", id: i.id, draft: { ...i } }); }
  function openNuevaPrestacion() { setModal({ type: "prestacion", mode: "new", draft: emptyPrestacion(instituciones) }); }
  function openEditarPrestacion(p) { setModal({ type: "prestacion", mode: "edit", id: p.id, draft: { ...p } }); }
  function openNuevaLocalidad() { setModal({ type: "localidad", mode: "new", draft: emptyLocalidad() }); }
  function openEditarLocalidad(l) { setModal({ type: "localidad", mode: "edit", id: l.id, draft: { ...l } }); }
  function openNuevaObraSocial() { setModal({ type: "obrasocial", mode: "new", draft: emptyObraSocial() }); }
  function openEditarObraSocial(o) { setModal({ type: "obrasocial", mode: "edit", id: o.id, draft: { ...o } }); }

  async function saveModal() {
    if (!modal) return;
    const { id: draftId, ...draftRest } = modal.draft;
    if (modal.type === "recorrido") {
      if (modal.mode === "new") {
        const newId = await saveItem("recorridos", null, { ...draftRest, chicos: [] });
        setSelectedId(newId);
      } else {
        await saveItem("recorridos", recorrido.id, { ...draftRest });
      }
    } else if (modal.type === "concurrente") {
      let chicoData = modal.draft;
      const manual = (chicoData.coordsManual || "").trim();
      const manualMatch = /^(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)$/.exec(manual);
      if (manualMatch) {
        chicoData = { ...chicoData, lat: parseFloat(manualMatch[1]), lng: parseFloat(manualMatch[2]) };
      } else if (chicoData.domicilio && chicoData.domicilio.trim()) {
        const coords = await geocodeAddress(chicoData.domicilio).catch(() => null);
        if (coords) chicoData = { ...chicoData, lat: coords.lat, lng: coords.lng };
      }
      const chicos = [...recorrido.chicos];
      if (modal.mode === "new") chicos.push(chicoData);
      else chicos[modal.index] = chicoData;
      const { id: rId, ...rRest } = recorrido;
      await saveItem("recorridos", recorrido.id, { ...rRest, chicos });
    } else if (modal.type === "usuario") {
      await saveItem("usuarios", modal.mode === "new" ? null : modal.id, draftRest);
    } else if (modal.type === "institucion") {
      await saveItem("instituciones", modal.mode === "new" ? null : modal.id, draftRest);
    } else if (modal.type === "prestacion") {
      await saveItem("prestaciones", modal.mode === "new" ? null : modal.id, draftRest);
    } else if (modal.type === "localidad") {
      await saveItem("localidades", modal.mode === "new" ? null : modal.id, draftRest);
    } else if (modal.type === "obrasocial") {
      await saveItem("obrasSociales", modal.mode === "new" ? null : modal.id, draftRest);
    }
    setModal(null);
  }

  async function deleteModal() {
    if (!modal) return;
    if (modal.type === "recorrido") {
      await deleteItem("recorridos", recorrido.id);
      setSelectedId((prev) => (recorridos[0]?.id === recorrido.id ? recorridos[1]?.id : recorridos[0]?.id));
    } else if (modal.type === "concurrente") {
      const chicos = recorrido.chicos.filter((_, i) => i !== modal.index);
      const { id: rId, ...rRest } = recorrido;
      await saveItem("recorridos", recorrido.id, { ...rRest, chicos });
    } else if (modal.type === "usuario") {
      await deleteItem("usuarios", modal.id);
    } else if (modal.type === "institucion") {
      await deleteItem("instituciones", modal.id);
    } else if (modal.type === "prestacion") {
      await deleteItem("prestaciones", modal.id);
    } else if (modal.type === "localidad") {
      await deleteItem("localidades", modal.id);
    } else if (modal.type === "obrasocial") {
      await deleteItem("obrasSociales", modal.id);
    }
    setModal(null);
  }

  function getExportRows() {
    const rows = [];
    recorridos
      .filter((r) => exportRecorridoIds.includes(r.id))
      .forEach((r) => {
        const kmDiarioR = r.kmIda * 2;
        const kmMensualR = kmDiarioR * diasPorMes;
        r.chicos
          .filter((c) => exportInstitucionIds.includes(c.institucionId))
          .forEach((c) => {
            const inst = instituciones.find((i) => i.id === c.institucionId);
            const base = {
              Recorrido: r.nombre,
              "Localidad recorrido": r.localidad,
              Concurrente: c.nombre,
              DNI: c.dni,
              Prestación: c.prestacion,
              "Obra social": c.obraSocial,
              "Localidad concurrente": c.localidad,
              Días: c.dias,
              Domicilio: c.domicilio,
              Institución: inst ? inst.nombre : "",
              "Distancia a institución (km)": c.kmDesde || 0,
              "Distancia diaria concurrente (km)": c.kmDesde ? c.kmDesde * 2 : 0,
              "Km ida recorrido": r.kmIda,
              "Km diario recorrido": kmDiarioR,
              [`Km mensual recorrido (${diasPorMes} días)`]: Math.round(kmMensualR),
            };
            if (exportIncluirTransporte) {
              base.Chofer = r.chofer;
              base.Auxiliar = r.auxiliar;
              base.Vehículo = r.vehiculo;
              base.Patente = r.patente;
            }
            rows.push(base);
          });
      });
    return rows;
  }

  function exportExcel() {
    const rows = getExportRows();
    if (rows.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Transporte");
    XLSX.writeFile(wb, `transporte_turno-${turno}.xlsx`);
  }

  function pdfEscape(text) {
    return String(text).replace(/—/g, "-").replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
  }

  function pdfTrunc(text, widthPt, size) {
    const avgCharW = size * 0.62;
    const maxChars = Math.max(3, Math.floor((widthPt - 6) / avgCharW));
    const t = String(text || "");
    return t.length <= maxChars ? t : t.slice(0, maxChars - 2) + "..";
  }

  function buildPdfBytes(pages) {
    const W = 595, H = 842;
    let nextId = 1;
    const catalogId = nextId++, pagesId = nextId++, fontRegId = nextId++, fontBoldId = nextId++;
    const objects = [];
    const kidsIds = [];

    pages.forEach((page) => {
      const pageId = nextId++, contentId = nextId++;
      kidsIds.push(pageId);
      let stream = "";
      (page.hlines || []).forEach((l) => {
        const y = (H - l.y).toFixed(1);
        stream += `${l.color} RG\n${l.lw} w\n${l.x1} ${y} m ${l.x2} ${y} l S\n`;
      });
      stream += "BT\n";
      (page.texts || []).forEach((t) => {
        const font = t.bold ? "/F2" : "/F1";
        const size = t.size || 9;
        const yPdf = (H - t.y).toFixed(1);
        stream += `${font} ${size} Tf\n1 0 0 1 ${t.x} ${yPdf} Tm\n(${pdfEscape(t.text)}) Tj\n`;
      });
      stream += "ET";
      objects.push({ id: contentId, body: `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream` });
      objects.push({
        id: pageId,
        body: `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /Font << /F1 ${fontRegId} 0 R /F2 ${fontBoldId} 0 R >> >> /Contents ${contentId} 0 R >>`,
      });
    });

    objects.push({ id: fontRegId, body: `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>` });
    objects.push({ id: fontBoldId, body: `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>` });
    objects.push({ id: pagesId, body: `<< /Type /Pages /Kids [${kidsIds.map((k) => k + " 0 R").join(" ")}] /Count ${kidsIds.length} >>` });
    objects.push({ id: catalogId, body: `<< /Type /Catalog /Pages ${pagesId} 0 R >>` });
    objects.sort((a, b) => a.id - b.id);

    let pdf = "%PDF-1.4\n";
    const offsets = {};
    objects.forEach((obj) => {
      offsets[obj.id] = pdf.length;
      pdf += `${obj.id} 0 obj\n${obj.body}\nendobj\n`;
    });
    const xrefStart = pdf.length;
    const totalObjs = objects.length + 1;
    pdf += `xref\n0 ${totalObjs}\n0000000000 65535 f \n`;
    for (let id = 1; id < nextId; id++) pdf += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
    pdf += `trailer\n<< /Size ${totalObjs} /Root ${catalogId} 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

    const bytes = new Uint8Array(pdf.length);
    for (let i = 0; i < pdf.length; i++) bytes[i] = pdf.charCodeAt(i) & 0xff;
    return bytes;
  }

  function exportPDF() {
    const GREEN_RGB = "0.247 0.424 0.318";
    const GRAY_RGB = "0.82 0.85 0.82";
    const marginX = 40, tableW = 515, top = 50, bottom = 792;
    const colDefs = [
      { key: "idx", label: "#", w: 14 },
      { key: "nombre", label: "Concurrente", w: 74 },
      { key: "dni", label: "DNI", w: 48 },
      { key: "institucion", label: "Institución", w: 58 },
      { key: "obraSocial", label: "O. social", w: 36 },
      { key: "km", label: "Km", w: 30 },
      { key: "dias", label: "Días", w: 82 },
      { key: "kmDiario", label: "Km/día", w: 38 },
      { key: "prestacion", label: "Prestación", w: 80 },
    ];
    let cursorX = marginX;
    const cols = colDefs.map((c) => {
      const withX = { ...c, x: cursorX };
      cursorX += c.w;
      return withX;
    });

    let pages = [{ texts: [], hlines: [] }];
    let y = top;
    const newPage = () => { pages.push({ texts: [], hlines: [] }); y = top; };
    const curPage = () => pages[pages.length - 1];
    const addText = (text, x, size = 9, bold = false) => curPage().texts.push({ text, x, y, size, bold });
    const addHLine = (color, lw) => curPage().hlines.push({ x1: marginX, x2: marginX + tableW, y, color, lw });
    const wrapText = (text, maxWidthPt, size) => {
      const maxChars = Math.max(10, Math.floor(maxWidthPt / (size * 0.52)));
      const words = String(text).split(" ");
      const lines = [];
      let current = "";
      words.forEach((w) => {
        const candidate = current ? current + " " + w : w;
        if (candidate.length > maxChars && current) {
          lines.push(current);
          current = w;
        } else {
          current = candidate;
        }
      });
      if (current) lines.push(current);
      return lines;
    };
    const addWrappedText = (text, x, maxWidthPt, size = 9, bold = false, lineH = 12) => {
      wrapText(text, maxWidthPt, size).forEach((line) => {
        if (y + lineH > bottom) newPage();
        addText(line, x, size, bold);
        y += lineH;
      });
    };
    const drawTableHeader = () => {
      cols.forEach((c) => curPage().texts.push({ text: c.label, x: c.x + 2, y: y + 11, size: 8, bold: true }));
      y += 16;
      addHLine(GREEN_RGB, 1);
      y += 1;
    };

    addText(`Transporte de concurrentes - Turno ${turno}`, marginX, 14, true);
    y += 24;

    recorridos.filter((r) => exportRecorridoIds.includes(r.id)).forEach((r) => {
      const chicos = r.chicos.filter((c) => exportInstitucionIds.includes(c.institucionId));
      if (chicos.length === 0) return;

      if (y + 60 > bottom) newPage();
      addText(`${r.nombre} - ${r.localidad}`, marginX, 12, true);
      y += 18;
      if (exportIncluirTransporte) {
        addWrappedText(`Chofer: ${r.chofer}   Auxiliar: ${r.auxiliar}   Vehiculo: ${r.vehiculo}   Patente: ${r.patente}`, marginX, tableW, 9, false, 12);
      }
      addText(`Km ida: ${r.kmIda.toFixed(1)}   Km diario (ida+vuelta): ${(r.kmIda * 2).toFixed(1)}   Km mensual (${diasPorMes} dias): ${Math.round(r.kmIda * 2 * diasPorMes)}`, marginX, 9);
      y += 16;

      drawTableHeader();

      chicos.forEach((c, i) => {
        if (y + 15 > bottom) {
          newPage();
          addText(`${r.nombre} (continuación)`, marginX, 10, true);
          y += 16;
          drawTableHeader();
        }
        const inst = instituciones.find((x) => x.id === c.institucionId);
        const rowY = y + 10;
        const values = {
          idx: String(i + 1),
          nombre: c.nombre,
          dni: c.dni,
          institucion: inst ? inst.nombre : "-",
          obraSocial: c.obraSocial,
          km: c.kmDesde ? c.kmDesde.toFixed(1) : "-",
          kmDiario: c.kmDesde ? (c.kmDesde * 2).toFixed(1) : "-",
          dias: c.dias,
          prestacion: c.prestacion,
        };
        cols.forEach((c2) => {
          curPage().texts.push({ text: pdfTrunc(values[c2.key], c2.w, 7.5), x: c2.x + 2, y: rowY, size: 7.5, bold: false });
        });
        y += 15;
        addHLine(GRAY_RGB, 0.5);
      });
      y += 14;
    });

    const bytes = buildPdfBytes(pages);
    const blob = new Blob([bytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `transporte_turno-${turno}.pdf`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  const d = modal?.draft;

  if (screen === "login") {
    return <Login usuarios={usuarios} onEnter={(r, matched) => { setRole(r); setLoggedInUsuario(matched); setScreen("landing"); }} />;
  }

  if (screen === "landing") {
    return <Landing onEnter={(t) => { setTurno(t); setScreen("app"); }} />;
  }

  if (!recorridosReady) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ background: "#F3F5F1" }}>
        <style>{FONT_IMPORT}</style>
        <div className="text-center">
          <img src={LOGO_URL} alt="Despertares" className="h-12 mx-auto object-contain opacity-70" />
          <p className="mt-4 text-[13px]" style={{ color: MUTED }}>Cargando datos...</p>
        </div>
      </div>
    );
  }

  if (role === "usuario") {
    const misRecorrido = recorridos.find((r) => r.id === loggedInUsuario?.recorridoId);
    return (
      <div className="min-h-screen w-full flex flex-col" style={{ background: "#F3F5F1", fontFamily: "Inter" }}>
        <style>{FONT_IMPORT}</style>
        <div className="flex items-center justify-between px-5 py-4 bg-white border-b border-[#E2E7E2]">
          <div className="flex items-center gap-2">
            <img src={LOGO_URL} alt="Despertares" className="h-7 w-auto object-contain" />
            <div>
              <div className="text-[14px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>{loggedInUsuario?.nombre}</div>
              <div className="text-[11px] flex items-center gap-1" style={{ color: MUTED }}>
                {loggedInUsuario?.rol} · {turno === "mañana" ? <Sun size={11} /> : <Moon size={11} />} Turno {turno}
              </div>
            </div>
          </div>
          <button onClick={() => { setLoggedInUsuario(null); setScreen("login"); }}
            className="flex items-center gap-1.5 text-[12px] rounded-full px-3 py-1.5 border" style={{ borderColor: "#E2E7E2", color: MUTED }}>
            <LogOut size={13} /> Salir
          </button>
        </div>

        {!misRecorrido ? (
          <div className="flex-1 flex items-center justify-center px-6 text-center">
            <p className="text-[13.5px]" style={{ color: MUTED }}>Todavía no tenés un recorrido asignado. Hablá con el administrador.</p>
          </div>
        ) : (
          <main className="flex-1 overflow-y-auto">
            <div className="max-w-[720px] mx-auto px-4 sm:px-8 py-6">
              <div className="text-[11px] uppercase tracking-wider" style={{ color: MUTED }}>{misRecorrido.localidad}</div>
              <h1 className="text-[22px] sm:text-[26px] mt-1" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>{misRecorrido.nombre}</h1>

              <div className="mt-5 grid grid-cols-2 sm:grid-cols-4 gap-3">
                {[
                  { icon: User, label: "Chofer", value: misRecorrido.chofer },
                  { icon: User, label: "Auxiliar", value: misRecorrido.auxiliar },
                  { icon: Bus, label: "Vehículo", value: misRecorrido.vehiculo || "-" },
                  { icon: CreditCard, label: "Patente", value: misRecorrido.patente },
                ].map((f, i) => (
                  <div key={i} className="rounded-xl bg-white border border-[#E2E7E2] px-3.5 py-3">
                    <div className="flex items-center gap-1.5" style={{ color: MUTED }}><f.icon size={12.5} /><span className="text-[10.5px] uppercase tracking-wide">{f.label}</span></div>
                    <div className="text-[13.5px] mt-1" style={{ color: INK, fontWeight: 500 }}>{f.value}</div>
                  </div>
                ))}
              </div>

              <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] p-5">
                <h3 className="text-[15px] mb-4" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Mapa del recorrido</h3>
                <RouteMap recorrido={misRecorrido} instituciones={instituciones} />
              </div>

              <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] overflow-hidden">
                <div className="px-5 py-4 border-b border-[#E2E7E2]">
                  <h3 className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Concurrentes ({misRecorrido.chicos.length})</h3>
                </div>
                <div className="divide-y divide-[#EDF0ED]">
                  {misRecorrido.chicos.map((c, i) => (
                    <div key={i} className="px-5 py-3.5">
                      <div className="flex items-center gap-1.5 text-[14px]" style={{ color: INK, fontWeight: 500 }}>
                        <span className="inline-flex items-center justify-center rounded-full text-[10px] w-4 h-4 shrink-0" style={{ background: "#EEF3EC", color: GREEN }}>{i + 1}</span>
                        {c.nombre}
                      </div>
                      <div className="mt-1 flex flex-wrap gap-x-4 gap-y-1 text-[12px]" style={{ color: "#5B6B63" }}>
                        {c.dias && <span className="flex items-center gap-1"><Calendar size={11} /> {c.dias}</span>}
                        {(c.domicilio || c.localidad) && <span className="flex items-center gap-1"><MapPin size={11} /> {[c.domicilio, c.localidad].filter(Boolean).join(" · ")}</span>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </main>
        )}
      </div>
    );
  }

  if (!recorrido) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ background: "#F3F5F1" }}>
        <style>{FONT_IMPORT}</style>
        <div className="text-center">
          <img src={LOGO_URL} alt="Despertares" className="h-12 mx-auto object-contain opacity-70" />
          <p className="mt-4 text-[13px]" style={{ color: MUTED }}>Todavía no hay recorridos cargados.</p>
        </div>
      </div>
    );
  }

  const navItems = [
    { id: "recorridos", label: "Recorridos", icon: Bus },
    { id: "importar", label: "Importar", icon: Upload },
    { id: "exportar", label: "Exportar", icon: Download },
    { id: "usuarios", label: "Usuarios", icon: Users },
    { id: "configuracion", label: "Datos generales", icon: Settings },
  ];

  return (
    <div className="min-h-screen w-full flex flex-col md:flex-row" style={{ background: "#F3F5F1", fontFamily: "Inter" }}>
      <style>{FONT_IMPORT}</style>

      {/* Barra superior mobile con hamburguesa */}
      <div className="md:hidden flex items-center justify-between px-4 py-3 bg-white border-b border-[#E2E7E2]">
        <div className="flex items-center gap-2">
          <img src={LOGO_URL} alt="Despertares" className="h-7 w-auto object-contain" />
          <span className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>
            {navItems.find((n) => n.id === section)?.label}
          </span>
        </div>
        <button onClick={() => setMobileNavOpen(true)} className="p-1.5">
          <Menu size={22} color={INK} />
        </button>
      </div>

      {/* Menú lateral: fijo a la izquierda en desktop */}
      <aside className="hidden md:flex md:w-[76px] shrink-0 border-r border-[#E2E7E2] bg-white flex-col items-center py-5">
        <img src={LOGO_URL} alt="Despertares" className="h-8 w-auto object-contain mb-4" />
        <div className="flex flex-col items-center gap-1">
          {navItems.map((n) => {
            const active = section === n.id;
            return (
              <button key={n.id} onClick={() => setSection(n.id)}
                className="flex flex-col items-center justify-center gap-0.5 rounded-xl w-14 h-14 shrink-0"
                style={{ background: active ? "#EEF3EC" : "transparent" }}>
                <n.icon size={17} color={active ? GREEN : MUTED} />
                <span className="text-[9.5px]" style={{ color: active ? GREEN : MUTED, fontWeight: active ? 600 : 400 }}>{n.label}</span>
              </button>
            );
          })}
        </div>
        <button onClick={() => { setLoggedInUsuario(null); setScreen("login"); }}
          className="flex flex-col items-center justify-center gap-0.5 rounded-xl w-14 h-14 shrink-0 mt-auto">
          <LogOut size={17} color={MUTED} />
          <span className="text-[9.5px]" style={{ color: MUTED }}>Salir</span>
        </button>
      </aside>

      {/* Drawer mobile */}
      {mobileNavOpen && (
        <div className="fixed inset-0 z-50 flex md:hidden">
          <div className="absolute inset-0 bg-black/30" onClick={() => setMobileNavOpen(false)} />
          <div className="relative w-72 max-w-[80%] bg-white h-full flex flex-col py-5 px-4">
            <div className="flex items-center justify-between px-1 mb-5">
              <div className="flex items-center gap-2">
                <img src={LOGO_URL} alt="Despertares" className="h-7 w-auto object-contain" />
                <span className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Menú</span>
              </div>
              <button onClick={() => setMobileNavOpen(false)}><X size={20} color={MUTED} /></button>
            </div>
            <div className="flex flex-col gap-1">
              {navItems.map((n) => {
                const active = section === n.id;
                return (
                  <button key={n.id} onClick={() => { setSection(n.id); setMobileNavOpen(false); }}
                    className="flex items-center gap-3 rounded-xl px-3 py-3 text-left"
                    style={{ background: active ? "#EEF3EC" : "transparent" }}>
                    <n.icon size={17} color={active ? GREEN : MUTED} />
                    <span className="text-[14px]" style={{ color: active ? "#2C4A34" : INK, fontWeight: active ? 600 : 400 }}>{n.label}</span>
                  </button>
                );
              })}
            </div>

            {section === "recorridos" && (
              <div className="mt-5 pt-4 border-t border-[#EDF0ED] flex-1 overflow-y-auto">
                <div className="text-[10.5px] uppercase tracking-wide px-3 mb-2" style={{ color: MUTED }}>Recorridos</div>
                <div className="flex flex-col gap-1">
                  {filtered.map((r) => {
                    const active = r.id === selectedId;
                    return (
                      <button key={r.id} onClick={() => { setSelectedId(r.id); setMobileNavOpen(false); }}
                        className="text-left rounded-lg px-3 py-2.5"
                        style={{ background: active ? "#EEF3EC" : "transparent" }}>
                        <div className="text-[13.5px]" style={{ fontWeight: 500, color: active ? "#2C4A34" : INK }}>{r.nombre}</div>
                        <div className="flex items-center gap-1.5 mt-0.5 text-[11px]" style={{ color: MUTED }}>
                          <MapPin size={10} /><span>{r.localidad}</span><span className="mx-0.5">·</span><span>{r.chicos.length} chicos</span>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            <button onClick={() => { setMobileNavOpen(false); setLoggedInUsuario(null); setScreen("login"); }}
              className="flex items-center gap-3 rounded-xl px-3 py-3 text-left mt-4">
              <LogOut size={17} color={MUTED} />
              <span className="text-[14px]" style={{ color: INK }}>Salir</span>
            </button>
          </div>
        </div>
      )}

      <div className="flex-1 flex flex-col md:flex-row min-w-0">
        {section === "recorridos" && (
          <>
            {recorridosPanelOpen ? (
              <aside className="hidden md:flex shrink-0 md:border-r border-[#E2E7E2] bg-white flex-col" style={{ width: isDesktop ? 280 : "100%" }}>
                <div className="px-5 pt-5 pb-3">
                  <div className="flex items-center justify-between">
                    <div>
                      <span className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Recorridos</span>
                      <div className="text-[11px] flex items-center gap-1 mt-0.5" style={{ color: MUTED }}>
                        {turno === "mañana" ? <Sun size={11} /> : <Moon size={11} />} Turno {turno}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <button onClick={() => setRecorridosPanelOpen(false)} title="Ocultar panel" className="p-1.5 rounded-lg" style={{ color: GREEN, background: "#EEF3EC" }}>
                        <Pin size={14} />
                      </button>
                      <button onClick={openNuevoRecorrido} className="flex items-center gap-1 text-[12px] rounded-lg px-2.5 py-1.5 text-white" style={{ background: GREEN }}>
                        <Plus size={13} /> Nuevo
                      </button>
                    </div>
                  </div>
                  <div className="mt-3 relative">
                    <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#8B978F]" />
                    <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar recorrido o localidad"
                      className="w-full text-[13px] rounded-lg border border-[#E2E7E2] bg-[#F9FAF8] pl-8 pr-3 py-2 outline-none focus:border-[#3F6C51]" />
                  </div>
                </div>
                <nav
                  className="flex-1 px-3 pb-3 md:pb-4"
                  style={{ display: "flex", flexDirection: isDesktop ? "column" : "row", gap: isDesktop ? 4 : 8, overflowY: isDesktop ? "auto" : "visible", overflowX: isDesktop ? "visible" : "auto" }}
                >
                  {filtered.map((r) => {
                    const active = r.id === selectedId;
                    return (
                      <button key={r.id} onClick={() => setSelectedId(r.id)}
                        className="text-left rounded-lg px-3 py-3 transition-colors"
                      style={{
                        background: active ? "#EEF3EC" : "transparent",
                        border: active ? "1px solid #C9DBC9" : "1px solid #EDF0ED",
                        flexShrink: 0,
                        width: isDesktop ? "100%" : 210,
                      }}>
                      <div className="flex items-center justify-between">
                        <span className="text-[13.5px]" style={{ fontWeight: 500, color: active ? "#2C4A34" : INK }}>{r.nombre}</span>
                        {isDesktop && <ChevronRight size={14} color={active ? GREEN : "#C7CFC7"} />}
                      </div>
                      <div className="flex items-center gap-1.5 mt-1 text-[11.5px]" style={{ color: MUTED }}>
                        <MapPin size={11} /><span>{r.localidad}</span><span className="mx-0.5">·</span><span>{r.chicos.length} chicos</span>
                      </div>
                    </button>
                  );
                })}
                </nav>
              </aside>
            ) : (
              <div className="hidden md:flex shrink-0 border-r border-[#E2E7E2] bg-white flex-col items-center pt-5">
                <button onClick={() => setRecorridosPanelOpen(true)} title="Mostrar panel de recorridos" className="p-2 rounded-lg" style={{ color: GREEN, background: "#EEF3EC" }}>
                  <PinOff size={14} />
                </button>
              </div>
            )}

            <main className="flex-1 overflow-y-auto">
              <div className="max-w-[820px] mx-auto px-4 sm:px-8 py-6 sm:py-8">
                <div className="flex flex-wrap items-start justify-between gap-2">
                  <div>
                    <div className="text-[11px] uppercase tracking-wider" style={{ color: MUTED }}>{recorrido.localidad}</div>
                    <h1 className="text-[22px] sm:text-[26px] mt-1" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>{recorrido.nombre}</h1>
                  </div>
                  <div className="flex items-center gap-2">
                    <button onClick={openEditarRecorrido} className="flex items-center gap-1.5 text-[12px] rounded-full px-3 py-1.5 border border-[#E2E7E2]" style={{ color: INK }}>
                      <Pencil size={12} /> Editar
                    </button>
                    <span className="text-[12px] rounded-full px-3 py-1.5 shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace", background: "#EEF3EC", color: GREEN }}>{recorrido.patente}</span>
                  </div>
                </div>

                <div className="mt-5 grid grid-cols-2 sm:grid-cols-4 gap-3">
                  {[
                    { icon: User, label: "Chofer", value: recorrido.chofer },
                    { icon: User, label: "Auxiliar", value: recorrido.auxiliar },
                    { icon: Bus, label: "Vehículo", value: recorrido.vehiculo },
                    { icon: CreditCard, label: "Patente", value: recorrido.patente },
                  ].map((f, i) => (
                    <div key={i} className="rounded-xl bg-white border border-[#E2E7E2] px-3.5 py-3">
                      <div className="flex items-center gap-1.5" style={{ color: MUTED }}><f.icon size={12.5} /><span className="text-[10.5px] uppercase tracking-wide">{f.label}</span></div>
                      <div className="text-[13.5px] mt-1" style={{ color: INK, fontWeight: 500 }}>{f.value}</div>
                    </div>
                  ))}
                </div>

                <div className="mt-6 flex flex-wrap gap-3 items-stretch">
                  <Stat icon={Milestone} label="Ida (recogida completa)" value={recorrido.kmIda.toFixed(1)} unit="km" />
                  <Stat icon={Milestone} label="Diario (ida + vuelta)" value={kmDiario.toFixed(1)} unit="km" accent={GREEN} />
                  <div className="flex-1 min-w-[180px] rounded-xl bg-white border border-[#E2E7E2] px-4 py-3">
                    <div className="flex items-center justify-between" style={{ color: MUTED }}>
                      <div className="flex items-center gap-1.5"><Calendar size={14} /><span className="text-[11px] tracking-wide uppercase">Mensual</span></div>
                      <input type="number" value={diasPorMes} onChange={(e) => setDiasPorMes(Number(e.target.value) || 0)}
                        className="w-12 text-[11px] text-center rounded border border-[#E2E7E2] py-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }} />
                    </div>
                    <div className="mt-1.5 flex items-baseline gap-1">
                      <span className="text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: OCHRE, fontWeight: 500 }}>{kmMensual.toFixed(0)}</span>
                      <span className="text-xs" style={{ color: MUTED }}>km · {diasPorMes} días hábiles</span>
                    </div>
                  </div>
                </div>

                <div className="mt-7 grid grid-cols-1 lg:grid-cols-[1.15fr_0.85fr] gap-5">
                  <div className="rounded-xl bg-white border border-[#E2E7E2] overflow-hidden">
                    <div className="px-5 py-4 border-b border-[#E2E7E2] flex items-center justify-between gap-2 flex-wrap">
                      <h3 className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Concurrentes del recorrido</h3>
                      <div className="flex items-center gap-2">
                        <button onClick={optimizarOrden} disabled={optimizando} className="flex items-center gap-1 text-[11.5px] rounded-full px-2.5 py-1.5 border" style={{ borderColor: "#E2E7E2", color: optimizando ? MUTED : GREEN }}>
                          <Route size={12} /> {optimizando ? "Optimizando..." : "Optimizar orden"}
                        </button>
                        <button onClick={openNuevoConcurrente} className="flex items-center gap-1 text-[11.5px] rounded-full px-2.5 py-1.5 text-white" style={{ background: GREEN }}>
                          <Plus size={12} /> Agregar
                        </button>
                      </div>
                    </div>
                    {optimizarMsg && (
                      <div className="px-5 py-2 text-[11.5px]" style={{ background: optimizarMsg.startsWith("No") ? "#FBEEEC" : "#EEF3EC", color: optimizarMsg.startsWith("No") ? "#B5533E" : GREEN }}>
                        {optimizarMsg}
                      </div>
                    )}
                    <div className="divide-y divide-[#EDF0ED]">
                      {recorrido.chicos.map((c, i) => {
                        const inst = instituciones.find((x) => x.id === c.institucionId);
                        const isOpen = expandedChico === i;
                        return (
                          <div key={i} className="px-5 py-3">
                            <button
                              onClick={() => setExpandedChico(isOpen ? null : i)}
                              className="w-full flex items-center justify-between gap-2 text-left"
                            >
                              <span className="text-[14px] flex items-center gap-1.5 min-w-0" style={{ color: INK, fontWeight: 500 }}>
                                <span className="inline-flex items-center justify-center rounded-full text-[10px] w-4 h-4 shrink-0" style={{ background: "#EEF3EC", color: GREEN }}>{i + 1}</span>
                                <span className="truncate">{c.nombre}</span>
                              </span>
                              <span className="flex items-center gap-2 shrink-0">
                                <span
                                  onClick={(e) => { e.stopPropagation(); setExpandedChico(isOpen ? null : i); }}
                                  className="text-[11px] rounded-full px-2.5 py-1 border"
                                  style={{ borderColor: "#E2E7E2", color: MUTED }}
                                >
                                  {isOpen ? "Ocultar" : "Ver datos"}
                                </span>
                                <ChevronRight size={14} color={MUTED} style={{ transform: isOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s" }} />
                              </span>
                            </button>

                            {isOpen && (
                              <div className="mt-3 pl-5">
                                <div className="flex items-center justify-between gap-2 mb-2">
                                  {c.dni && <span className="text-[11.5px]" style={{ color: MUTED, fontFamily: "'IBM Plex Mono', monospace" }}>DNI {c.dni}</span>}
                                  <button onClick={() => openEditarConcurrente(i)} className="flex items-center gap-1 text-[11.5px] ml-auto" style={{ color: GREEN }}>
                                    <Pencil size={12} /> Editar
                                  </button>
                                </div>
                                <div className="flex flex-wrap gap-x-4 gap-y-1.5 text-[12px]" style={{ color: "#5B6B63" }}>
                                  {c.prestacion && <span className="flex items-center gap-1"><HeartPulse size={12} /> {c.prestacion}</span>}
                                  {c.obraSocial && <span className="flex items-center gap-1">Obra social: {c.obraSocial}</span>}
                                  {c.dias && <span className="flex items-center gap-1"><Calendar size={12} /> {c.dias}</span>}
                                </div>
                                {(c.domicilio || c.localidad) && (
                                  <div className="mt-1.5 flex items-center gap-1 text-[12px]" style={{ color: MUTED }}>
                                    <MapPin size={12} /> {[c.domicilio, c.localidad].filter(Boolean).join(" · ")}
                                  </div>
                                )}
                                <div className="mt-1.5 flex items-center gap-1 text-[11.5px]" style={{ color: GREEN }}>
                                  <Building2 size={12} /> {inst ? inst.nombre : "Sin institución asignada"}
                                </div>
                                {typeof c.kmDesde === "number" && c.kmDesde > 0 && (
                                  <div className="mt-1.5 flex items-center gap-1 text-[11.5px]" style={{ color: OCHRE }}>
                                    <Milestone size={12} /> {c.kmDesde.toFixed(1)} km desde {inst ? inst.nombre : "la institución"}
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })}
                      {recorrido.chicos.length === 0 && (
                        <div className="px-5 py-8 text-center text-[12.5px]" style={{ color: MUTED }}>Todavía no hay concurrentes en este recorrido.</div>
                      )}
                    </div>
                  </div>

                  <div className="rounded-xl bg-white border border-[#E2E7E2] p-5">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-[15px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>
                        {view === "diagrama" ? "Recorrido esquemático" : "Mapa del recorrido"}
                      </h3>
                      <div className="flex items-center rounded-full border border-[#E2E7E2] p-0.5">
                        <button onClick={() => setView("diagrama")} className="p-1.5 rounded-full" style={{ background: view === "diagrama" ? "#EEF3EC" : "transparent" }}>
                          <List size={13} color={view === "diagrama" ? GREEN : MUTED} />
                        </button>
                        <button onClick={() => setView("mapa")} className="p-1.5 rounded-full" style={{ background: view === "mapa" ? "#EEF3EC" : "transparent" }}>
                          <MapIcon size={13} color={view === "mapa" ? GREEN : MUTED} />
                        </button>
                      </div>
                    </div>
                    {view === "diagrama" ? <RouteDiagram recorrido={recorrido} instituciones={instituciones} /> : <RouteMap recorrido={recorrido} instituciones={instituciones} />}
                  </div>
                </div>
              </div>
            </main>
          </>
        )}

        {section === "importar" && (
          <main className="flex-1 overflow-y-auto">
            <div className="max-w-[640px] mx-auto px-4 sm:px-8 py-6 sm:py-8">
              <h1 className="text-[22px] sm:text-[26px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Importar desde Excel</h1>
              <p className="mt-1 text-[13px]" style={{ color: MUTED }}>
                Subí una planilla con bloques por "TRANSPORTE:" (como la que ya cargamos) para traer todos los recorridos y concurrentes de una vez.
              </p>

              <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] p-8 text-center">
                <input type="file" accept=".xlsx,.xls" id="import-file-input" className="hidden" onChange={handleImportFile} />
                <label htmlFor="import-file-input" className="inline-flex items-center gap-2 rounded-full px-5 py-2.5 text-white text-[13px] cursor-pointer" style={{ background: GREEN }}>
                  <Upload size={15} /> Elegir archivo Excel
                </label>
                <p className="mt-3 text-[11.5px]" style={{ color: MUTED }}>Formato .xlsx, mismo diseño que la planilla de turnos por transporte</p>
              </div>

              {importError && (
                <div className="mt-4 rounded-lg px-4 py-3 text-[12.5px]" style={{ background: "#FBEEEC", color: "#B5533E" }}>{importError}</div>
              )}

              {importPreview && (
                <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] p-5">
                  <div className="flex items-center gap-1.5 mb-1">
                    <FileSpreadsheet size={15} color={GREEN} />
                    <span className="text-[14px]" style={{ color: INK, fontWeight: 500 }}>
                      Se encontraron {importPreview.recorridos.length} recorridos y {importPreview.recorridos.reduce((a, r) => a + r.chicos.length, 0)} concurrentes
                    </span>
                  </div>
                  <div className="mt-3 max-h-[260px] overflow-y-auto divide-y divide-[#EDF0ED]">
                    {importPreview.recorridos.map((r) => (
                      <div key={r.id} className="py-2 flex items-center justify-between text-[12.5px]">
                        <span style={{ color: INK }}>{r.nombre}</span>
                        <span style={{ color: MUTED }}>{r.chicos.length} chicos · {r.localidad}</span>
                      </div>
                    ))}
                  </div>
                  {importPreview.prestacionesNuevas.length > 0 && (
                    <div className="mt-4 rounded-lg px-3 py-2.5 text-[12px]" style={{ background: "#FBF1E6", color: "#8A5A24" }}>
                      {importPreview.prestacionesNuevas.length} prestación(es) nueva(s) sin sede asignada: {importPreview.prestacionesNuevas.map((p) => p.nombre).join(", ")}.
                      Las vas a poder asignar a una sede en Configuración → Prestaciones después de importar.
                    </div>
                  )}
                  {importPreview.sinCoordsPropias > 0 && (
                    <div className="mt-4 rounded-lg px-3 py-2.5 text-[12px]" style={{ background: "#FBEEEC", color: "#B5533E" }}>
                      {importPreview.sinCoordsPropias} concurrente(s) no tenían coordenadas legibles en la planilla — quedaron con la ubicación de la institución en su lugar. Revisá esa columna en el Excel (a veces el signo menos se pega distinto) o cargalas a mano después con "Coordenadas manuales".
                    </div>
                  )}
                  <div className="mt-5 pt-4 border-t border-[#EDF0ED] flex items-center justify-between gap-3">
                    <p className="text-[11.5px]" style={{ color: MUTED }}>Esto reemplaza los recorridos que ya tenés cargados (las sedes no se tocan).</p>
                    <button onClick={confirmImport} className="shrink-0 text-[13px] rounded-full px-4 py-2 text-white" style={{ background: GREEN }}>
                      Reemplazar recorridos
                    </button>
                  </div>
                </div>
              )}
            </div>
          </main>
        )}

        {section === "exportar" && (
          <main className="flex-1 overflow-y-auto">
            <div className="max-w-[640px] mx-auto px-4 sm:px-8 py-6 sm:py-8">
              <h1 className="text-[22px] sm:text-[26px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Exportar</h1>
              <p className="mt-1 text-[13px]" style={{ color: MUTED }}>Elegí qué recorridos e instituciones incluir, y exportá</p>

              <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] p-5">
                <div className="flex items-center justify-between mb-2.5">
                  <span className="text-[13px]" style={{ color: INK, fontWeight: 500 }}>Recorridos</span>
                  <div className="flex gap-3 text-[11.5px]" style={{ color: GREEN }}>
                    <button onClick={() => setExportRecorridoIds(recorridos.map((r) => r.id))}>Todos</button>
                    <button onClick={() => setExportRecorridoIds([])}>Ninguno</button>
                  </div>
                </div>
                <div className="space-y-1.5">
                  {recorridos.map((r) => {
                    const checked = exportRecorridoIds.includes(r.id);
                    return (
                      <label key={r.id} className="flex items-center gap-2.5 text-[13.5px] px-1 py-1.5 cursor-pointer" style={{ color: INK }}>
                        <input type="checkbox" checked={checked}
                          onChange={(e) => setExportRecorridoIds((prev) => e.target.checked ? [...prev, r.id] : prev.filter((id) => id !== r.id))}
                          className="accent-[#3F6C51] w-4 h-4" />
                        {r.nombre} <span style={{ color: MUTED }}>· {r.chicos.length} chicos</span>
                      </label>
                    );
                  })}
                </div>

                <div className="flex items-center justify-between mt-5 mb-2.5">
                  <span className="text-[13px]" style={{ color: INK, fontWeight: 500 }}>Instituciones</span>
                  <div className="flex gap-3 text-[11.5px]" style={{ color: GREEN }}>
                    <button onClick={() => setExportInstitucionIds(instituciones.map((i) => i.id))}>Todas</button>
                    <button onClick={() => setExportInstitucionIds([])}>Ninguna</button>
                  </div>
                </div>
                <div className="space-y-1.5">
                  {instituciones.map((i) => {
                    const checked = exportInstitucionIds.includes(i.id);
                    return (
                      <label key={i.id} className="flex items-center gap-2.5 text-[13.5px] px-1 py-1.5 cursor-pointer" style={{ color: INK }}>
                        <input type="checkbox" checked={checked}
                          onChange={(e) => setExportInstitucionIds((prev) => e.target.checked ? [...prev, i.id] : prev.filter((id) => id !== i.id))}
                          className="accent-[#3F6C51] w-4 h-4" />
                        {i.nombre}
                      </label>
                    );
                  })}
                </div>

                <label className="flex items-center gap-2.5 text-[13.5px] px-1 py-1.5 mt-5 pt-4 border-t border-[#EDF0ED] cursor-pointer" style={{ color: INK }}>
                  <input type="checkbox" checked={exportIncluirTransporte} onChange={(e) => setExportIncluirTransporte(e.target.checked)} className="accent-[#3F6C51] w-4 h-4" />
                  Incluir datos del transporte (chofer, auxiliar, vehículo, patente)
                </label>

                <div className="mt-5 pt-4 border-t border-[#EDF0ED] flex items-center justify-between">
                  <span className="text-[12px]" style={{ color: MUTED }}>{getExportRows().length} concurrentes coinciden con el filtro</span>
                </div>

                <div className="mt-4 flex flex-col sm:flex-row gap-2.5">
                  <button onClick={exportExcel} className="flex-1 flex items-center justify-center gap-2 text-[13px] rounded-lg px-4 py-2.5 text-white" style={{ background: GREEN }}>
                    <FileSpreadsheet size={15} /> Exportar a Excel
                  </button>
                  <button onClick={exportPDF}
                    className="flex-1 flex items-center justify-center gap-2 text-[13px] rounded-lg px-4 py-2.5 border border-[#E2E7E2]" style={{ color: INK }}>
                    <Printer size={15} /> Exportar a PDF
                  </button>
                </div>
              </div>
              <p className="mt-4 text-[11.5px] leading-relaxed" style={{ color: MUTED }}>
                Ambos botones descargan el archivo directo (Excel y PDF), sin pasar por el diálogo de impresión del navegador.
              </p>
            </div>
          </main>
        )}

        {section === "usuarios" && (
          <main className="flex-1 overflow-y-auto">
            <div className="max-w-[720px] mx-auto px-4 sm:px-8 py-6 sm:py-8">
              <div className="flex items-start justify-between gap-2">
                <div>
                  <h1 className="text-[22px] sm:text-[26px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Usuarios del sistema</h1>
                  <p className="mt-1 text-[13px]" style={{ color: MUTED }}>Choferes y auxiliares con acceso para ver su recorrido asignado y guiarse por el mapa</p>
                </div>
                <button onClick={openNuevoUsuario} className="flex items-center gap-1.5 text-[12.5px] rounded-full px-3.5 py-2 text-white shrink-0" style={{ background: GREEN }}>
                  <UserPlus size={14} /> Nuevo usuario
                </button>
              </div>

              <div className="mt-6 rounded-xl bg-white border border-[#E2E7E2] overflow-hidden divide-y divide-[#EDF0ED]">
                {usuarios.map((u) => {
                  const rec = recorridos.find((r) => r.id === u.recorridoId);
                  return (
                    <div key={u.id} className="px-5 py-4 flex items-center justify-between gap-3">
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "#EEF3EC" }}>
                          <User size={15} color={GREEN} />
                        </div>
                        <div className="min-w-0">
                          <div className="text-[14px] truncate" style={{ color: INK, fontWeight: 500 }}>{u.nombre}</div>
                          <div className="text-[11.5px]" style={{ color: MUTED }}>{u.rol} · {rec ? rec.nombre : "Sin recorrido asignado"} · usuario: {u.usuario}</div>
                        </div>
                      </div>
                      <button onClick={() => openEditarUsuario(u)} className="shrink-0"><Pencil size={14} color={MUTED} /></button>
                    </div>
                  );
                })}
                {usuarios.length === 0 && <div className="px-5 py-8 text-center text-[12.5px]" style={{ color: MUTED }}>Todavía no hay usuarios cargados.</div>}
              </div>
              <p className="mt-5 text-[11.5px] leading-relaxed" style={{ color: MUTED }}>
                Cada chofer o auxiliar va a poder entrar con su usuario y contraseña y ver únicamente el recorrido que
                tiene asignado (con el mapa y la lista de chicos), sin acceder a los datos del resto de la institución.
                Útil, por ejemplo, si un chofer suplente reemplaza a otro y necesita guiarse.
              </p>
            </div>
          </main>
        )}

        {section === "configuracion" && (
          <main className="flex-1 overflow-y-auto">
            <div className="max-w-[640px] mx-auto px-4 sm:px-8 py-6 sm:py-8">
              <h1 className="text-[22px] sm:text-[26px]" style={{ fontFamily: "Fraunces", color: INK, fontWeight: 500 }}>Datos generales</h1>
              <p className="mt-1 text-[13px]" style={{ color: MUTED }}>
                Catálogos compartidos por todos los formularios: instituciones, prestaciones, obras sociales y localidades.
              </p>

              <div className="mt-5 flex flex-wrap gap-1.5">
                {[
                  { id: "instituciones", label: "Instituciones" },
                  { id: "prestaciones", label: "Prestaciones" },
                  { id: "obrasociales", label: "Obras sociales" },
                  { id: "localidades", label: "Localidades" },
                ].map((t) => {
                  const active = datosTab === t.id;
                  return (
                    <button key={t.id} onClick={() => setDatosTab(t.id)}
                      className="text-[12.5px] rounded-full px-3.5 py-1.5"
                      style={{ background: active ? GREEN : "#EEF3EC", color: active ? "#FFFFFF" : GREEN }}>
                      {t.label}
                    </button>
                  );
                })}
              </div>

              {datosTab === "instituciones" && (
                <div>
                  <div className="flex items-center justify-between mt-6 mb-2">
                    <p className="text-[13px]" style={{ color: MUTED }}>Sedes a las que puede pertenecer cada concurrente</p>
                    <button onClick={openNuevaInstitucion} className="flex items-center gap-1.5 text-[12.5px] rounded-full px-3.5 py-2 text-white shrink-0" style={{ background: GREEN }}>
                      <Plus size={14} /> Nueva sede
                    </button>
                  </div>
                  <div className="rounded-xl bg-white border border-[#E2E7E2] overflow-hidden divide-y divide-[#EDF0ED]">
                    {instituciones.map((i) => (
                      <div key={i.id} className="px-5 py-4 flex items-center justify-between gap-3">
                        <div className="flex items-center gap-3 min-w-0">
                          <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "#EEF3EC" }}>
                            <Building2 size={15} color={GREEN} />
                          </div>
                          <div className="min-w-0">
                            <div className="text-[14px] truncate" style={{ color: INK, fontWeight: 500 }}>{i.nombre}</div>
                            <div className="text-[11.5px] truncate" style={{ color: MUTED }}>{i.direccion}</div>
                          </div>
                        </div>
                        <button onClick={() => openEditarInstitucion(i)} className="shrink-0"><Pencil size={14} color={MUTED} /></button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {datosTab === "prestaciones" && (
                <div>
                  <div className="flex items-center justify-between mt-6 mb-2">
                    <p className="text-[13px]" style={{ color: MUTED }}>Programas (SAIE, SET, CET, Centro de Día, Hogar Permanente...) y su sede</p>
                    <button onClick={openNuevaPrestacion} className="flex items-center gap-1.5 text-[12.5px] rounded-full px-3.5 py-2 text-white shrink-0" style={{ background: GREEN }}>
                      <Plus size={14} /> Nueva prestación
                    </button>
                  </div>
                  <div className="rounded-xl bg-white border border-[#E2E7E2] overflow-hidden divide-y divide-[#EDF0ED]">
                    {prestaciones.map((p) => {
                      const inst = instituciones.find((i) => i.id === p.institucionId);
                      return (
                        <div key={p.id} className="px-5 py-4 flex items-center justify-between gap-3">
                          <div className="flex items-center gap-3 min-w-0">
                            <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "#FBF1E6" }}>
                              <HeartPulse size={15} color={OCHRE} />
                            </div>
                            <div className="min-w-0">
                              <div className="text-[14px] truncate" style={{ color: INK, fontWeight: 500 }}>{p.nombre}</div>
                              <div className="text-[11.5px] truncate flex items-center gap-1" style={{ color: MUTED }}>
                                <Building2 size={11} /> {inst ? inst.nombre : "Sin sede asignada"}
                              </div>
                            </div>
                          </div>
                          <button onClick={() => openEditarPrestacion(p)} className="shrink-0"><Pencil size={14} color={MUTED} /></button>
                        </div>
                      );
                    })}
                    {prestaciones.length === 0 && (
                      <div className="px-5 py-8 text-center text-[12.5px]" style={{ color: MUTED }}>Todavía no hay prestaciones cargadas.</div>
                    )}
                  </div>
                </div>
              )}

              {datosTab === "obrasociales" && (
                <div>
                  <div className="flex items-center justify-between mt-6 mb-2">
                    <p className="text-[13px]" style={{ color: MUTED }}>Obras sociales disponibles para elegir en cada concurrente</p>
                    <button onClick={openNuevaObraSocial} className="flex items-center gap-1.5 text-[12.5px] rounded-full px-3.5 py-2 text-white shrink-0" style={{ background: GREEN }}>
                      <Plus size={14} /> Nueva obra social
                    </button>
                  </div>
                  <div className="rounded-xl bg-white border border-[#E2E7E2] overflow-hidden divide-y divide-[#EDF0ED]">
                    {obrasSociales.map((o) => (
                      <div key={o.id} className="px-5 py-4 flex items-center justify-between gap-3">
                        <span className="text-[14px]" style={{ color: INK, fontWeight: 500 }}>{o.nombre}</span>
                        <button onClick={() => openEditarObraSocial(o)} className="shrink-0"><Pencil size={14} color={MUTED} /></button>
                      </div>
                    ))}
                    {obrasSociales.length === 0 && (
                      <div className="px-5 py-8 text-center text-[12.5px]" style={{ color: MUTED }}>Todavía no hay obras sociales cargadas.</div>
                    )}
                  </div>
                </div>
              )}

              {datosTab === "localidades" && (
                <div>
                  <div className="flex items-center justify-between mt-6 mb-2">
                    <p className="text-[13px]" style={{ color: MUTED }}>Localidad y provincia de origen de cada concurrente</p>
                    <button onClick={openNuevaLocalidad} className="flex items-center gap-1.5 text-[12.5px] rounded-full px-3.5 py-2 text-white shrink-0" style={{ background: GREEN }}>
                      <Plus size={14} /> Nueva localidad
                    </button>
                  </div>
                  <div className="rounded-xl bg-white border border-[#E2E7E2] overflow-hidden divide-y divide-[#EDF0ED]">
                    {localidades.map((l) => (
                      <div key={l.id} className="px-5 py-4 flex items-center justify-between gap-3">
                        <div className="flex items-center gap-1.5 text-[14px]" style={{ color: INK, fontWeight: 500 }}>
                          <MapPin size={13} color={GREEN} /> {l.nombre} <span style={{ color: MUTED, fontWeight: 400 }}>· {l.provincia}</span>
                        </div>
                        <button onClick={() => openEditarLocalidad(l)} className="shrink-0"><Pencil size={14} color={MUTED} /></button>
                      </div>
                    ))}
                    {localidades.length === 0 && (
                      <div className="px-5 py-8 text-center text-[12.5px]" style={{ color: MUTED }}>Todavía no hay localidades cargadas.</div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </main>
        )}
      </div>

      {modal && modal.type === "recorrido" && (
        <Modal title={modal.mode === "new" ? "Nuevo recorrido" : "Editar recorrido"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre del recorrido" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <Field label="Localidad" value={d.localidad} onChange={(v) => setModal({ ...modal, draft: { ...d, localidad: v } })} />
          <Field label="Chofer" value={d.chofer} onChange={(v) => setModal({ ...modal, draft: { ...d, chofer: v } })} />
          <Field label="Auxiliar" value={d.auxiliar} onChange={(v) => setModal({ ...modal, draft: { ...d, auxiliar: v } })} />
          <Field label="Vehículo" value={d.vehiculo} onChange={(v) => setModal({ ...modal, draft: { ...d, vehiculo: v } })} />
          <Field label="Patente" value={d.patente} onChange={(v) => setModal({ ...modal, draft: { ...d, patente: v } })} />
          <Field label="Km ida (recogida completa)" type="number" value={d.kmIda} onChange={(v) => setModal({ ...modal, draft: { ...d, kmIda: v } })} />
        </Modal>
      )}

      {modal && modal.type === "concurrente" && (
        <Modal title={modal.mode === "new" ? "Nuevo concurrente" : "Editar concurrente"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre y apellido" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <Field label="DNI (opcional)" value={d.dni} onChange={(v) => setModal({ ...modal, draft: { ...d, dni: v } })} />
          <SelectField label="Prestación" value={d.prestacion} onChange={(v) => {
            const pres = prestaciones.find((p) => p.nombre === v);
            setModal({ ...modal, draft: { ...d, prestacion: v, institucionId: pres ? pres.institucionId : d.institucionId } });
          }} options={prestaciones.map((p) => ({ value: p.nombre, label: p.nombre }))} />
          <SelectField label="Institución (sede)" value={d.institucionId || ""} onChange={(v) => setModal({ ...modal, draft: { ...d, institucionId: v } })}
            options={instituciones.map((i) => ({ value: i.id, label: i.nombre }))} />
          <p className="text-[11px]" style={{ color: MUTED }}>Se sugiere sola según la prestación, pero se puede cambiar acá si el concurrente pasó a otra sede.</p>
          <SelectField label="Obra social (opcional)" value={d.obraSocial} onChange={(v) => setModal({ ...modal, draft: { ...d, obraSocial: v } })}
            options={[{ value: "", label: "Sin especificar" }, ...obrasSociales.map((o) => ({ value: o.nombre, label: o.nombre }))]} />
          <SelectField label="Localidad" value={d.localidad} onChange={(v) => setModal({ ...modal, draft: { ...d, localidad: v } })}
            options={[{ value: "", label: "Sin especificar" }, ...localidades.map((l) => ({ value: l.nombre, label: `${l.nombre}, ${l.provincia}` }))]} />
          <Field label="Días de concurrencia" value={d.dias} onChange={(v) => setModal({ ...modal, draft: { ...d, dias: v } })} />
          <Field label="Domicilio" value={d.domicilio} onChange={(v) => setModal({ ...modal, draft: { ...d, domicilio: v } })} />
          {typeof d.lat === "number" && typeof d.lng === "number" && (
            <p className="text-[11.5px] flex items-center gap-1" style={{ color: GREEN }}>
              <MapPin size={12} /> Coordenadas guardadas: {d.lat.toFixed(6)}, {d.lng.toFixed(6)}
            </p>
          )}
          <Field label="Coordenadas manuales (lat, lng)" value={d.coordsManual || ""} onChange={(v) => setModal({ ...modal, draft: { ...d, coordsManual: v } })} />
          <p className="text-[11px]" style={{ color: MUTED }}>
            Dejar vacío para mantener las coordenadas guardadas de arriba. Completar solo para reemplazarlas (ej. si el domicilio no tiene calle y número, zona rural). Formato: -26.144265, -59.599022 — si se completa, tiene prioridad sobre el domicilio.
          </p>
          <Field label="Distancia a la institución (km)" type="number" value={d.kmDesde || 0} onChange={(v) => setModal({ ...modal, draft: { ...d, kmDesde: v } })} />
          <p className="text-[11px]" style={{ color: MUTED }}>En la versión final, el domicilio se geocodifica solo al guardar. Obra social, localidad, institución y prestación se cargan y editan desde Datos generales.</p>
        </Modal>
      )}

      {modal && modal.type === "usuario" && (
        <Modal title={modal.mode === "new" ? "Nuevo usuario" : "Editar usuario"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre y apellido" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <SelectField label="Rol" value={d.rol} onChange={(v) => setModal({ ...modal, draft: { ...d, rol: v } })}
            options={[{ value: "Chofer", label: "Chofer" }, { value: "Auxiliar", label: "Auxiliar" }]} />
          <SelectField label="Recorrido asignado" value={d.recorridoId} onChange={(v) => setModal({ ...modal, draft: { ...d, recorridoId: v } })}
            options={recorridos.map((r) => ({ value: r.id, label: r.nombre }))} />
          <Field label="Usuario" value={d.usuario} onChange={(v) => setModal({ ...modal, draft: { ...d, usuario: v } })} />
          <Field label="Contraseña" type="password" value={d.password || ""} onChange={(v) => setModal({ ...modal, draft: { ...d, password: v } })} />
          <p className="text-[11px]" style={{ color: MUTED }}>Con este usuario y contraseña, el chofer/auxiliar va a poder entrar y ver únicamente el recorrido asignado.</p>
        </Modal>
      )}

      {modal && modal.type === "institucion" && (
        <Modal title={modal.mode === "new" ? "Nueva sede" : "Editar sede"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre de la sede" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <Field label="Dirección" value={d.direccion} onChange={(v) => setModal({ ...modal, draft: { ...d, direccion: v } })} />
          <p className="text-[11px]" style={{ color: MUTED }}>En la versión final, la dirección se geocodifica sola al guardar.</p>
        </Modal>
      )}

      {modal && modal.type === "prestacion" && (
        <Modal title={modal.mode === "new" ? "Nueva prestación" : "Editar prestación"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre de la prestación" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <SelectField label="Sede a la que pertenece" value={d.institucionId} onChange={(v) => setModal({ ...modal, draft: { ...d, institucionId: v } })}
            options={instituciones.map((i) => ({ value: i.id, label: i.nombre }))} />
        </Modal>
      )}

      {modal && modal.type === "localidad" && (
        <Modal title={modal.mode === "new" ? "Nueva localidad" : "Editar localidad"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Localidad" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
          <Field label="Provincia" value={d.provincia} onChange={(v) => setModal({ ...modal, draft: { ...d, provincia: v } })} />
        </Modal>
      )}

      {modal && modal.type === "obrasocial" && (
        <Modal title={modal.mode === "new" ? "Nueva obra social" : "Editar obra social"} onClose={() => setModal(null)} onSave={saveModal} onDelete={modal.mode === "edit" ? deleteModal : undefined}>
          <Field label="Nombre de la obra social" value={d.nombre} onChange={(v) => setModal({ ...modal, draft: { ...d, nombre: v } })} />
        </Modal>
      )}
    </div>
  );
}
