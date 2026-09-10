const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  HeadingLevel,
  AlignmentType,
  PageBreak
} = require('docx');

console.log('Generando documento Word con Normas APA 7ma edición...');

// Configuración de estilo APA 7
const FONT_FAMILY = 'Times New Roman';
const COLOR_TEXT = '000000';
const COLOR_HEADER = '1B365D'; // Azul institucional elegante

// Bordes estilo APA para tablas (sólo bordes horizontales superior e inferior, y debajo de cabecera)
const tableBorderNone = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const tableBorderThin = { style: BorderStyle.SINGLE, size: 8, color: '333333' };
const tableBorderThick = { style: BorderStyle.SINGLE, size: 16, color: '000000' };

function p(text, options = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, font: FONT_FAMILY, size: options.size || 24, bold: options.bold, italics: options.italics, color: options.color || COLOR_TEXT })];
  return new Paragraph({
    children: runs,
    alignment: options.alignment || AlignmentType.LEFT,
    spacing: {
      line: options.lineSpacing || 360, // 1.5 spacing
      before: options.spaceBefore || 120,
      after: options.spaceAfter || 120
    }
  });
}

function heading1(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: FONT_FAMILY,
        size: 28, // 14pt
        bold: true,
        color: COLOR_HEADER
      })
    ],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 360 }
  });
}

function heading2(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: FONT_FAMILY,
        size: 26, // 13pt
        bold: true,
        color: '2C3E50'
      })
    ],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120, line: 360 }
  });
}

function tableAPAHeaderCell(text, widthPercent) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    borders: {
      top: tableBorderThick,
      bottom: tableBorderThin,
      left: tableBorderNone,
      right: tableBorderNone
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text, bold: true, font: FONT_FAMILY, size: 22 })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 80, after: 80 }
      })
    ]
  });
}

function tableAPACell(text, widthPercent, isLastRow = false, isBold = false) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    borders: {
      top: tableBorderNone,
      bottom: isLastRow ? tableBorderThick : tableBorderNone,
      left: tableBorderNone,
      right: tableBorderNone
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text, bold: isBold, font: FONT_FAMILY, size: 20 })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 60, after: 60 }
      })
    ]
  });
}

// -------------------------------------------------------------
// CONSTRUCCIÓN DEL DOCUMENTO
// -------------------------------------------------------------

const children = [];

// ==========================================
// 1. PORTADA ESTUDIANTIL SEGÚN NORMAS APA 7
// ==========================================
children.push(
  new Paragraph({
    children: [
      new TextRun({
        text: 'UNIVERSIDAD MARIANO GÁLVEZ DE GUATEMALA',
        bold: true,
        font: FONT_FAMILY,
        size: 28,
        color: COLOR_HEADER
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 80 }
  }),
  new Paragraph({
    children: [
      new TextRun({
        text: 'FACULTAD DE INGENIERÍA EN SISTEMAS DE INFORMACIÓN Y CIENCIAS DE LA COMPUTACIÓN\nLICENCIATURA EN INGENIERÍA EN SISTEMAS DE INFORMACIÓN\nCURSO: DESARROLLO WEB — DÉCIMO SEMESTRE',
        bold: true,
        font: FONT_FAMILY,
        size: 22,
        color: '333333'
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 }
  }),
  new Paragraph({
    children: [
      new TextRun({
        text: 'SISTEMA DE ESCAPE ROOM DIGITAL COOPERATIVO EN TIEMPO REAL: ARQUITECTURA EN CAPAS CON CONTROLADORES, SERVICIOS DE NEGOCIO, DTOs Y PERSISTENCIA RELACIONAL EN NESTJS Y POSTGRESQL',
        bold: true,
        font: FONT_FAMILY,
        size: 28,
        color: '000000'
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 600 }
  }),
  new Paragraph({
    children: [
      new TextRun({
        text: 'INFORME TÉCNICO Y ESPECIFICACIÓN FORMAL DE SOFTWARE\nBAJO DIRECTRICES DE LAS NORMAS APA (7MA EDICIÓN)',
        italics: true,
        font: FONT_FAMILY,
        size: 22
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 700 }
  }),
  new Paragraph({
    children: [
      new TextRun({ text: 'Autor: ', bold: true, font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Estudiante de Ingeniería en Sistemas de Información\n', font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Catedrático: ', bold: true, font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Catedrático del Curso de Desarrollo Web\n', font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Institución: ', bold: true, font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Universidad Mariano Gálvez de Guatemala\n', font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Fecha: ', bold: true, font: FONT_FAMILY, size: 24 }),
      new TextRun({ text: 'Septiembre de 2026', font: FONT_FAMILY, size: 24 })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 }
  }),
  new Paragraph({ children: [new PageBreak()] })
);

// ==========================================
// 2. RESUMEN Y ABSTRACT
// ==========================================
children.push(
  heading1('Resumen'),
  p(
    'El presente informe técnico-académico describe el análisis, diseño e implementación integral de la infraestructura backend para una plataforma de Escape Room Digital Cooperativo Multijugador. La solución tecnológica fue concebida bajo el paradigma de arquitectura empresarial en capas y desacoplamiento estricto, utilizando el framework NestJS 11 sobre Node.js y el lenguaje tipado TypeScript. El núcleo del sistema implementa rigurosamente el patrón arquitectónico compuesto por Controladores (Controllers) para la gestión de solicitudes HTTP/REST y Swagger/OpenAPI, Servicios de Negocio (Business Services) para la gobernanza de las reglas y penalizaciones lógicas del juego, y Objetos de Transferencia de Datos (DTOs) con tuberías de validación automáticas en tiempo de ejecución sustentadas en class-validator y class-transformer. Asimismo, el sistema incorpora un canal de comunicación bidireccional de baja latencia mediante WebSockets (@WebSocketGateway) para la orquestación sincrónica de acertijos grupales, y una capa de persistencia relacional impulsada por el motor PostgreSQL 18.3 a través del ORM tipado Prisma. El documento demuestra el consumo efectivo de la base de datos para la auditoría histórica de partidas y el cálculo dinámico de la tabla de clasificación (Leaderboard), cumpliendo con los estándares de calidad de software y documentación formal bajo las Normas APA (7ma edición).'
  ),
  p([
    new TextRun({ text: 'Palabras clave: ', bold: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'NestJS, Angular, Data Transfer Object (DTO), WebSockets, PostgreSQL, Prisma ORM, Arquitectura en Capas, Escape Room Digital.', italics: true, font: FONT_FAMILY, size: 24 })
  ], { spaceAfter: 400 }),

  heading1('Abstract'),
  p(
    'This academic and technical report details the analysis, architectural design, and software implementation of the backend infrastructure for a real-time Collaborative Digital Escape Room platform. The software solution was built upon an enterprise layered architecture pattern utilizing the NestJS 11 framework powered by Node.js and TypeScript. The core system strictly implements an architectural decoupling comprising Controllers for handling RESTful endpoints and Swagger/OpenAPI documentation, Business Services to encapsulate game logic rules and time-penalty algorithms, and Data Transfer Objects (DTOs) with automatic runtime validation pipelines powered by class-validator and class-transformer. Furthermore, the application provides a bi-directional, low-latency communication layer using WebSockets (@WebSocketGateway) for multi-agent game state synchronization, and a relational persistence engine based on PostgreSQL 18.3 orchestrated via Prisma ORM. Live persistence and operational queries are demonstrated through game auditing and real-time leaderboard generation, strictly adhering to modern software engineering standards and formal APA 7th Edition guidelines.'
  ),
  p([
    new TextRun({ text: 'Keywords: ', bold: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'NestJS, Angular, Data Transfer Object (DTO), WebSockets, PostgreSQL, Prisma ORM, Layered Architecture, Digital Escape Room.', italics: true, font: FONT_FAMILY, size: 24 })
  ]),
  new Paragraph({ children: [new PageBreak()] })
);

// ==========================================
// 3. INTRODUCCIÓN Y PLANTEAMIENTO DEL PROBLEMA
// ==========================================
children.push(
  heading1('1. Introducción'),
  p(
    'El auge de las aplicaciones interactivas en la web moderna exige una transición de arquitecturas monolíticas desacopladas hacia ecosistemas altamente tipados, tolerantes a fallos y capaces de coordinar estados asíncronos y transacciones relacionales en tiempo real. En el contexto pedagógico y formativo de la Universidad Mariano Gálvez de Guatemala, el desarrollo de software a nivel de décimo semestre requiere la puesta en práctica de patrones de diseño avanzados que garanticen mantenibilidad, seguridad en los datos de entrada y una clara separación de responsabilidades (Separation of Concerns).'
  ),
  p(
    'El proyecto Escape Room Digital emerge como un entorno colaborativo donde múltiples usuarios resuelven enigmas de lógica, combinatoria, secuencias y deducción en salas privadas con tiempo límite. Para respaldar esta experiencia sin fisuras de red ni vulnerabilidades de inyección, este proyecto implementa un backend robusto basado en el framework NestJS 11 y una base de datos relacional PostgreSQL 18.3 gestionada mediante Prisma ORM. El presente informe documenta detalladamente el diseño arquitectónico, el modelado de la base de datos, el flujo de procesamiento de peticiones mediante DTOs, las reglas de negocio codificadas en servicios y los resultados empíricos de las pruebas de integración y consumo efectivo de la base de datos.'
  ),

  heading1('2. Planteamiento del Problema, Justificación y Alcance'),
  heading2('2.1 Planteamiento del Problema'),
  p(
    'En el desarrollo de plataformas lúdicas y colaborativas en tiempo real, uno de los desafíos más críticos radica en la dispersión de la lógica de negocio y la ausencia de validación formal en los puntos de entrada (endpoints). Cuando un servidor procesa datos mal estructurados, parámetros faltantes o tipos incompatibles provenientes de clientes web móviles o de escritorio, se generan comportamientos impredecibles, caídas imprevistas del servidor (crashes), o inconsistencias de estado en partidas concurrentes. Adicionalmente, el almacenamiento de sesiones y puntuaciones en memoria volátil provoca la pérdida irrevocable de las métricas de desempeño cuando el servidor reinicia o escala horizontalmente, imposibilitando la auditoría del rendimiento de los jugadores y la generación de registros históricos fidedignos.'
  ),

  heading2('2.2 Justificación'),
  p(
    'La adopción de una arquitectura en capas fundamentada en NestJS, inspirada en la arquitectura hexagonal y la inyección de dependencias de Angular, soluciona estos problemas al establecer barreras defensivas estrictas. El uso obligatorio de Data Transfer Objects (DTOs) junto a interceptores de validación global (ValidationPipe) garantiza que ninguna carga útil (payload) corrupta o no autorizada alcance la lógica de dominio. Asimismo, la integración de una base de datos relacional robusta como PostgreSQL mediante Prisma ORM otorga seguridad de tipos de extremo a extremo (End-to-End Type Safety), garantizando que las partidas finalizadas, tiempos de resolución y clasificaciones de honor persistan de forma confiable, auditable y de alto rendimiento.'
  ),

  heading2('2.3 Alcance y Delimitación'),
  p(
    'El alcance del presente desarrollo abarca: 1) Creación de una API RESTful autodocumentada mediante la especificación OpenAPI (Swagger UI). 2) Definición estricta de DTOs para creación de salas, unión de agentes, validación de enigmas y almacenamiento de misiones. 3) Incorporación de servicios de negocio desacoplados con validaciones y reglas lógicas (penalizaciones de -15 segundos, desbloqueo secuencial de sectores, temporizadores y resolución de enigmas). 4) Persistencia relacional en PostgreSQL bajo el esquema escaperoom mediante Prisma ORM para el registro de misiones y la consulta del Leaderboard. 5) Coexistencia con un WebSocket Gateway (@WebSocketGateway) para la emisión de eventos de sala en tiempo real hacia clientes desarrollados en Angular.'
  ),

  heading1('3. Objetivos'),
  heading2('3.1 Objetivo General'),
  p(
    'Desarrollar una infraestructura de backend empresarial para un sistema de Escape Room Digital Cooperativo, utilizando el framework NestJS y el motor relacional PostgreSQL, fundamentada en un diseño estricto de tres capas con Controladores, Servicios de Negocio, DTOs con validaciones automáticas y consumo persistente a base de datos.'
  ),
  heading2('3.2 Objetivos Específicos'),
  p('1. Diseñar e implementar Controladores REST que expongan los puntos de acceso del sistema bajo estándares HTTP y documentación interactiva Swagger.'),
  p('2. Definir Data Transfer Objects (DTOs) robustos decorados con reglas de class-validator, rechazando de manera automática cualquier carga útil no válida con códigos HTTP 400 Bad Request.'),
  p('3. Encapsular la lógica de dominio y las penalizaciones de tiempo en Servicios de Negocio (Business Services), protegiendo las reglas de la aplicación contra dependencias externas.'),
  p('4. Diseñar e implementar el modelo de persistencia relacional en PostgreSQL empleando Prisma Client, garantizando la inserción, consulta y ordenamiento de misiones e historiales de juego.'),
  p('5. Ejecutar una batería de pruebas de integración automatizadas que certifique el funcionamiento continuo del ciclo petición-validación-servicio-persistencia.')
);

// ==========================================
// 4. MARCO TEÓRICO
// ==========================================
children.push(
  heading1('4. Marco Teórico y Conceptual'),
  heading2('4.1 Arquitectura en Tres Capas y Principios SOLID'),
  p(
    'El diseño de software empresarial se basa en la separación de responsabilidades propuesta por Martin (2018). En este modelo, la capa de presentación recibe las peticiones, desempaqueta parámetros y delega el flujo de trabajo sin ejecutar lógica de cálculo. La capa de negocio aplica las políticas del sistema, validaciones semánticas, cálculos de puntajes y flujos de juego. Finalmente, la capa de acceso a datos aísla el motor de base de datos relacional de la lógica de aplicación, permitiendo migraciones transparentes.'
  ),

  heading2('4.2 Framework NestJS 11 y Modularidad'),
  p(
    'NestJS es un framework progresivo para Node.js desarrollado con TypeScript. Aplica principios de programación orientada a objetos (POO), programación funcional (FP) y programación reactiva funcional (FRP). Implementa un contenedor de Inversión de Control (IoC) e Inyección de Dependencias (DI) análogo al de Spring Boot en Java o Angular en el frontend, lo cual facilita la modularidad, el desacoplamiento y las pruebas unitarias (Myśliwiec, 2024).'
  ),

  heading2('4.3 Data Transfer Objects (DTO) y Validación Declarativa'),
  p(
    'Un DTO es un objeto que transporta datos entre subsistemas sin portar lógica de negocio (Fowler, 2002). En la implementación de NestJS, se utilizan decoradores de class-validator (tales como @IsString, @IsInt, @Min, @Max, @Length). El ValidationPipe global inspecciona cada objeto entrante antes de que llegue al controlador. Si alguna restricción se infringe, intercepta la petición y responde automáticamente con un payload estandarizado que describe con precisión los campos erróneos con código HTTP 400 Bad Request.'
  ),

  heading2('4.4 Persistencia Relacional con PostgreSQL 18 y Prisma ORM'),
  p(
    'PostgreSQL es un sistema gestor de bases de datos relacional y orientado a objetos de código abierto de reconocida estabilidad y rendimiento en transacciones concurrentes (The PostgreSQL Global Development Group, 2024). Prisma es un ORM de última generación para Node.js y TypeScript que sustituye el mapeo manual SQL tradicional por un esquema declarativo unificado (schema.prisma). Prisma genera un cliente fuertemente tipado en tiempo de compilación (PrismaClient), erradicando discrepancias entre las estructuras de datos de la aplicación y las tablas de la base de datos (Prisma Data Inc., 2024).'
  )
);

// ==========================================
// 5. ARQUITECTURA DEL SISTEMA Y CAPAS
// ==========================================
children.push(
  heading1('5. Especificación y Diseño de la Arquitectura del Sistema'),
  heading2('5.1 Capa de Controladores (EscapeController)'),
  p(
    'El controlador REST EscapeController gestiona los puntos de entrada HTTP (/api/health, /api/scenarios, /api/rooms, /api/missions, etc.). Su responsabilidad exclusiva es validar el contrato del DTO mediante pipes, transformar la respuesta en JSON y asignar el código de estado HTTP adecuado. Cuenta además con anotaciones @ApiOperation y @ApiResponse para generar la documentación Swagger interactiva en /api/docs.'
  ),

  heading2('5.2 Capa de Lógica de Negocio (EscapeService y DatabaseService)'),
  p(
    'Los servicios encapsulan las reglas críticas del juego: 1) Generación de códigos únicos alfanuméricos de sala de 4 caracteres. 2) Control de capacidad (máximo 4 jugadores por sala). 3) Evaluación de enigmas y penalización automática de 15 segundos al temporizador de la misión ante respuestas erróneas. 4) Desbloqueo progresivo de sectores temáticos con bonificación de 1,000 puntos. 5) Persistencia automática de partidas concluidas en PostgreSQL.'
  ),

  heading2('5.3 Capa de DTOs y Validación Automática'),
  p(
    'Se definieron DTOs especializados: CreateRoomDto (valida nombre de anfitrión de 2 a 25 caracteres), JoinRoomDto (valida código de sala con expresión regular exacta de 4 caracteres alfanuméricos), ValidatePuzzleDto (verifica tipo de puzzle y solución), y SaveMissionDto (persiste datos finales de la partida en PostgreSQL con validaciones numéricas y enumeraciones de estado).'
  ),

  heading2('5.4 Capa de Comunicación en Tiempo Real (WebSocket Gateway)'),
  p(
    'El gateway EscapeGateway expone eventos Socket.IO en el puerto 3002 (@SubscribeMessage) para coordinar las acciones simultáneas de los agentes, permitiendo la sincronización de cables eléctricos, presiones de válvulas neumáticas, chat táctico y cronómetro grupal sin recargar la página web.'
  )
);

// ==========================================
// 6. DISEÑO DE BASE DE DATOS Y TABLA 1
// ==========================================
children.push(
  heading1('6. Diseño y Modelado de la Base de Datos'),
  p(
    'La base de datos relacional opera sobre el motor PostgreSQL 18.3, alojada localmente bajo el esquema físico escaperoom. La tabla central escape_missions almacena el histórico de partidas finalizadas y alimenta la tabla de líderes (Leaderboard).'
  ),
  p([
    new TextRun({ text: 'Tabla 1\n', bold: true, font: FONT_FAMILY, size: 22 }),
    new TextRun({ text: 'Estructura de la entidad de persistencia EscapeMission en PostgreSQL', italics: true, font: FONT_FAMILY, size: 22 })
  ], { spaceAfter: 80 }),

  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          tableAPAHeaderCell('Nombre del Campo', 22),
          tableAPAHeaderCell('Tipo de Dato', 18),
          tableAPAHeaderCell('Longitud / Rango', 18),
          tableAPAHeaderCell('Nulo', 10),
          tableAPAHeaderCell('Propósito / Restricción', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('id', 22, false, true),
          tableAPACell('INTEGER', 18),
          tableAPACell('32 bits', 18),
          tableAPACell('No', 10),
          tableAPACell('Llave primaria autoincremental (PRIMARY KEY).', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('roomCode', 22, false, true),
          tableAPACell('VARCHAR', 18),
          tableAPACell('10 caracteres', 18),
          tableAPACell('No', 10),
          tableAPACell('Código único de la sala de escape (room_code).', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('scenarioId', 22, false, true),
          tableAPACell('VARCHAR', 18),
          tableAPACell('50 caracteres', 18),
          tableAPACell('No', 10),
          tableAPACell('Clave técnica del escenario temático.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('scenarioName', 22, false, true),
          tableAPACell('VARCHAR', 18),
          tableAPACell('150 caracteres', 18),
          tableAPACell('No', 10),
          tableAPACell('Nombre legible del escenario de juego.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('status', 22, false, true),
          tableAPACell('VARCHAR', 18),
          tableAPACell('30 caracteres', 18),
          tableAPACell('No', 10),
          tableAPACell('Estado de misión (GAME_WON, GAME_OVER).', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('teamScore', 22, false, true),
          tableAPACell('INTEGER', 18),
          tableAPACell('>= 0', 18),
          tableAPACell('No', 10),
          tableAPACell('Puntaje acumulado por enigmas resueltos.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('timeRemaining', 22, false, true),
          tableAPACell('INTEGER', 18),
          tableAPACell('>= 0 seg', 18),
          tableAPACell('No', 10),
          tableAPACell('Segundos remanentes en el cronómetro al escapar.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('playersCount', 22, false, true),
          tableAPACell('INTEGER', 18),
          tableAPACell('1 a 4', 18),
          tableAPACell('No', 10),
          tableAPACell('Cantidad de agentes en el equipo.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('playersNames', 22, false, true),
          tableAPACell('TEXT', 18),
          tableAPACell('Variable', 18),
          tableAPACell('No', 10),
          tableAPACell('Nombres concatenados de los jugadores.', 32)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('completedAt', 22, true, true),
          tableAPACell('TIMESTAMPTZ', 18, true),
          tableAPACell('Microsegundos', 18, true),
          tableAPACell('No', 10, true),
          tableAPACell('Marca de tiempo UTC (CURRENT_TIMESTAMP).', 32, true)
        ]
      })
    ]
  }),
  p([
    new TextRun({ text: 'Nota. ', italics: true, font: FONT_FAMILY, size: 20 }),
    new TextRun({ text: 'La entidad está mapeada bajo el esquema escaperoom en PostgreSQL 18.3 mediante Prisma ORM bajo la tabla física escape_missions. Los índices aceleran el ordenamiento descendente por teamScore y timeRemaining.', font: FONT_FAMILY, size: 20 })
  ], { spaceBefore: 80, spaceAfter: 240 })
);

// ==========================================
// 7. CATÁLOGO DE ENDPOINTS Y TABLA 2
// ==========================================
children.push(
  heading1('7. Catálogo de Endpoints de la API REST y Contratos de Datos'),
  p(
    'A continuación se detallan los servicios web RESTful expuestos por el backend de Escape Room, incluyendo los métodos HTTP, contratos de datos (DTOs) y códigos de estado semánticos:'
  ),
  p([
    new TextRun({ text: 'Tabla 2\n', bold: true, font: FONT_FAMILY, size: 22 }),
    new TextRun({ text: 'Catálogo de Endpoints de la API RESTful de Escape Room', italics: true, font: FONT_FAMILY, size: 22 })
  ], { spaceAfter: 80 }),

  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          tableAPAHeaderCell('Método', 12),
          tableAPAHeaderCell('Ruta / Endpoint', 28),
          tableAPAHeaderCell('DTO / Parámetro', 22),
          tableAPAHeaderCell('Código HTTP', 14),
          tableAPAHeaderCell('Descripción y Regla de Negocio', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('GET', 12, false, true),
          tableAPACell('/api/health', 28),
          tableAPACell('Ninguno', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Comprueba estado del servidor y PostgreSQL.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('GET', 12, false, true),
          tableAPACell('/api/database/status', 28),
          tableAPACell('Ninguno', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Telemetría técnica del esquema y tabla en BD.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('GET', 12, false, true),
          tableAPACell('/api/scenarios', 28),
          tableAPACell('Ninguno', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Catálogo de 3 escenarios temáticos de juego.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('POST', 12, false, true),
          tableAPACell('/api/rooms', 28),
          tableAPACell('CreateRoomDto', 22),
          tableAPACell('201 Created', 14),
          tableAPACell('Genera sala con código de 4 caracteres.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('POST', 12, false, true),
          tableAPACell('/api/rooms/join', 28),
          tableAPACell('JoinRoomDto', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Unión de agente con validación de código.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('POST', 12, false, true),
          tableAPACell('/api/rooms/validate', 28),
          tableAPACell('ValidatePuzzleDto', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Evalúa respuesta y aplica penalización -15s.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('POST', 12, false, true),
          tableAPACell('/api/missions', 28),
          tableAPACell('SaveMissionDto', 22),
          tableAPACell('201 Created', 14),
          tableAPACell('Consumo a BD: Inserta partida en PostgreSQL.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('GET', 12, false, true),
          tableAPACell('/api/missions/leaderboard', 28),
          tableAPACell('limit (opcional)', 22),
          tableAPACell('200 OK', 14),
          tableAPACell('Consumo a BD: Ranking ordenado por puntaje.', 24)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('GET', 12, true, true),
          tableAPACell('/api/missions/history', 28, true),
          tableAPACell('limit (opcional)', 22, true),
          tableAPACell('200 OK', 14, true),
          tableAPACell('Consumo a BD: Historial cronológico de partidas.', 24, true)
        ]
      })
    ]
  }),
  p([
    new TextRun({ text: 'Nota. ', italics: true, font: FONT_FAMILY, size: 20 }),
    new TextRun({ text: 'Todas las rutas de escritura aplican ValidationPipe estricto (whitelist: true, forbidNonWhitelisted: true), rechazando cualquier parámetro malicioso o sobrante.', font: FONT_FAMILY, size: 20 })
  ], { spaceBefore: 80, spaceAfter: 240 })
);

// ==========================================
// 8. PRUEBAS AUTOMATIZADAS Y TABLA 3
// ==========================================
children.push(
  heading1('8. Verificación, Ejecución y Pruebas del Backend'),
  p(
    'Para certificar el funcionamiento de los componentes y la persistencia en base de datos, se ejecutó una suite de pruebas automatizada (test_api.ps1) contra el servidor en ejecución en el puerto 3002 y la base de datos PostgreSQL 18.3.'
  ),
  p([
    new TextRun({ text: 'Tabla 3\n', bold: true, font: FONT_FAMILY, size: 22 }),
    new TextRun({ text: 'Matriz de Resultados de las Pruebas de Integración y Consumo a Base de Datos', italics: true, font: FONT_FAMILY, size: 22 })
  ], { spaceAfter: 80 }),

  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          tableAPAHeaderCell('No.', 8),
          tableAPAHeaderCell('Caso de Prueba', 24),
          tableAPAHeaderCell('Entrada de Prueba', 26),
          tableAPAHeaderCell('Resultado Obtenido', 26),
          tableAPAHeaderCell('Estado', 16)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('1', 8),
          tableAPACell('Comprobación de Salud', 24),
          tableAPACell('GET /api/health', 26),
          tableAPACell('Servidor online, PostgreSQL CONNECTED', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('2', 8),
          tableAPACell('Telemetría PostgreSQL', 24),
          tableAPACell('GET /api/database/status', 26),
          tableAPACell('Esquema escaperoom y tabla verificados', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('3', 8),
          tableAPACell('Catálogo de Escenarios', 24),
          tableAPACell('GET /api/scenarios', 26),
          tableAPACell('3 escenarios temáticos obtenidos', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('4', 8),
          tableAPACell('Creación de Sala (DTO)', 24),
          tableAPACell('{ hostName: "Comandante Alpha" }', 26),
          tableAPACell('Sala generada con código de 4 caracteres', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('5', 8),
          tableAPACell('Unión de Jugador (DTO)', 24),
          tableAPACell('{ roomCode: "RV3B", playerName: "Beta" }', 26),
          tableAPACell('Agente incorporado con rol especialista', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('6', 8),
          tableAPACell('Rechazo de DTO Inválido', 24),
          tableAPACell('{ hostName: "X" } (longitud < 2)', 26),
          tableAPACell('HTTP 400 Bad Request por ValidationPipe', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('7', 8),
          tableAPACell('Inserción en PostgreSQL', 24),
          tableAPACell('POST /api/missions con Score 4850', 26),
          tableAPACell('Insertado con ID autoincremental en BD', 26),
          tableAPACell('Aprobado', 16, false, true)
        ]
      }),
      new TableRow({
        children: [
          tableAPACell('8', 8, true),
          tableAPACell('Consulta de Leaderboard', 24, true),
          tableAPACell('GET /api/missions/leaderboard', 26, true),
          tableAPACell('Ranking devuelto en vivo desde PostgreSQL', 26, true),
          tableAPACell('Aprobado', 16, true, true)
        ]
      })
    ]
  }),
  p([
    new TextRun({ text: 'Nota. ', italics: true, font: FONT_FAMILY, size: 20 }),
    new TextRun({ text: 'Todas las pruebas fueron ejecutadas sobre entorno Windows con Node.js v22 y PostgreSQL 18.3 local.', font: FONT_FAMILY, size: 20 })
  ], { spaceBefore: 80, spaceAfter: 240 })
);

// ==========================================
// 9. CONCLUSIONES Y RECOMENDACIONES
// ==========================================
children.push(
  heading1('9. Conclusiones y Recomendaciones'),
  heading2('9.1 Conclusiones'),
  p(
    '1. La arquitectura modular desacoplada de NestJS permite construir servicios backend de alta fidelidad técnica, donde la lógica de negocio se encuentra estrictamente protegida por capas de controladores y DTOs fuertemente tipados.'
  ),
  p(
    '2. El uso de class-validator en conjunto con el ValidationPipe global de NestJS elimina de raíz las vulnerabilidades originadas por cargas útiles maliciosas o malformadas, garantizando que el núcleo de la aplicación opere únicamente con datos confiables.'
  ),
  p(
    '3. La integración de Prisma ORM sobre PostgreSQL 18.3 demostró ser altamente eficiente para registrar misiones y calcular el Leaderboard en tiempo real, garantizando la consistencia y seguridad transaccional de los datos de juego.'
  ),
  p(
    '4. El ecosistema coordinado de NestJS 11 y Angular 22 ofrece una sinergia natural al compartir la misma filosofía de inyección de dependencias, tipado estricto en TypeScript y reactividad, resultando idóneo para proyectos de desarrollo web de décimo semestre.'
  ),

  heading2('9.2 Recomendaciones'),
  p(
    '1. Para despliegues en entornos de producción, se sugiere alojar la base de datos PostgreSQL en servicios gestionados en la nube con réplicas de solo lectura para soportar altas tasas de lectura en la tabla de clasificación.'
  ),
  p(
    '2. Implementar una capa de almacenamiento en caché en memoria con Redis para mitigar la sobrecarga de consultas recurrentes al Leaderboard en momentos de alta concurrencia de partidas.'
  ),
  p(
    '3. Incorporar autenticación mediante tokens JWT (JSON Web Tokens) con Passport en NestJS para asociar la identidad criptográfica de los usuarios a las salas creadas y puntuaciones registradas.'
  )
);

// ==========================================
// 10. REFERENCIAS BIBLIOGRÁFICAS (APA 7)
// ==========================================
children.push(
  heading1('10. Referencias Bibliográficas'),
  p([
    new TextRun({ text: 'Fowler, M. (2002). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Patterns of Enterprise Application Architecture. ', italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Addison-Wesley Professional.', font: FONT_FAMILY, size: 24 })
  ]),
  p([
    new TextRun({ text: 'Martin, R. C. (2018). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: "Clean Architecture: A Craftsman's Guide to Software Structure and Design. ", italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Prentice Hall.', font: FONT_FAMILY, size: 24 })
  ]),
  p([
    new TextRun({ text: 'Myśliwiec, K. (2024). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'NestJS: A progressive Node.js framework for building efficient, reliable and scalable server-side applications. ', italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'NestJS Documentation. https://docs.nestjs.com', font: FONT_FAMILY, size: 24 })
  ]),
  p([
    new TextRun({ text: 'OpenAPI Initiative. (2024). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'OpenAPI Specification v3.1.0. ', italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Linux Foundation. https://spec.openapis.org/oas/v3.1.0', font: FONT_FAMILY, size: 24 })
  ]),
  p([
    new TextRun({ text: 'Prisma Data Inc. (2024). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Prisma: Next-generation ORM for Node.js & TypeScript. ', italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'Prisma Documentation. https://www.prisma.io/docs', font: FONT_FAMILY, size: 24 })
  ]),
  p([
    new TextRun({ text: 'The PostgreSQL Global Development Group. (2024). ', font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: "PostgreSQL 18 Documentation: The world's most advanced open source database. ", italics: true, font: FONT_FAMILY, size: 24 }),
    new TextRun({ text: 'PostgreSQL Documentation. https://www.postgresql.org/docs/', font: FONT_FAMILY, size: 24 })
  ])
);

// Crear Documento
const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          margin: {
            top: 1440,    // 1 pulgada = 2.54 cm = 1440 twips (Normas APA)
            right: 1440,
            bottom: 1440,
            left: 1440
          }
        }
      },
      children: children
    }
  ]
});

// Guardar archivos .docx
async function generate() {
  const buffer = await Packer.toBuffer(doc);
  
  const destPath1 = path.join(__dirname, 'DOCUMENTACION_ESCAPEROOM_NORMAS_APA.docx');
  const destPath2 = path.join(__dirname, '..', 'DOCUMENTACION_ESCAPEROOM_NORMAS_APA.docx');
  
  fs.writeFileSync(destPath1, buffer);
  console.log(`✅ Documento Word generado en: ${destPath1}`);
  
  fs.writeFileSync(destPath2, buffer);
  console.log(`✅ Copia generada en la raíz: ${destPath2}`);
}

generate().catch(err => {
  console.error('❌ Error al generar documento Word:', err);
});
