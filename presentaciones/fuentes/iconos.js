// Genera los iconos PNG de la presentación P2R (blanco y azul marino) a partir de react-icons.
// Uso: NODE_PATH=<carpeta con node_modules> node iconos.js <carpeta de salida>
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const path = require('path');

const ICONOS = {
  industria: 'FaIndustry', caja: 'FaBoxOpen', demanda: 'FaChartLine', balanza: 'FaBalanceScale',
  cubos: 'FaCubes', calendario: 'FaCalendarAlt', play: 'FaPlay', panel: 'FaTachometerAlt',
  equipo: 'FaUsers', intercambio: 'FaExchangeAlt', correo: 'FaEnvelope', robot: 'FaRobot',
  bandera: 'FaFlag', usuario: 'FaUserTie', engranajes: 'FaCogs', almacen: 'FaWarehouse',
  diana: 'FaBullseye', check: 'FaCheck', euro: 'FaEuroSign', reloj: 'FaClock', ruta: 'FaRoute',
};
const COLORES = { b: '#FFFFFF', n: '#00263A' };
const salida = process.argv[2] || '.';

(async () => {
  for (const [nombre, comp] of Object.entries(ICONOS)) {
    for (const [suf, color] of Object.entries(COLORES)) {
      const svg = renderToStaticMarkup(React.createElement(fa[comp], { color, size: 256 }));
      await sharp(Buffer.from(svg)).resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png().toFile(path.join(salida, `${nombre}_${suf}.png`));
    }
  }
  console.log(Object.keys(ICONOS).length * 2, 'iconos');
})();
