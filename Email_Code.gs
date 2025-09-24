// Función actualizada para generar items del checklist con lógica completa
function generarItemChecklist(tipo, estado, datos) {
  let html = '';
  
  switch(tipo) {
    case 'ingles':
      if (estado === 'Cumple (B2)' || estado === 'Cumple (C1)' || estado === 'Cumple') {
        // Versión COMPLETADA - Verde
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #E8F5E9; border: 1px solid #4CAF50;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">✅</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #2E7D32; font-size: 15px;">Nivel de inglés: ${estado}</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → ¡Excelente! Este requisito está completo.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      } else {
        // Versión PENDIENTE - Rojo
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #ffffff; border: 1px solid #dee2e6;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">❌</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #DC3545; font-size: 15px;">Nivel de inglés: ${estado || 'No cumple'}</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → Requisito indispensable para titularte. Acredita tu nivel antes del cierre del semestre.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      }
      break;
      
    case 'servicio_social':
      const horasSS = datos.horas || 0;
      const horasFaltantes = Math.max(0, 480 - horasSS);
      
      if (horasFaltantes === 0) {
        // Versión COMPLETADA - Verde
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #E8F5E9; border: 1px solid #4CAF50;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">✅</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #2E7D32; font-size: 15px;">Servicio Social: Completo (${horasSS} horas)</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → ¡Felicidades! Has completado tus horas de servicio social.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      } else {
        // Versión PENDIENTE - Rojo
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #ffffff; border: 1px solid #dee2e6;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">❌</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #DC3545; font-size: 15px;">Servicio Social: ${horasSS}/480 horas</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → Te faltan <strong>${horasFaltantes} horas</strong> por completar<br>
                        <strong>¿Dónde revisar tus horas?</strong><br>
                        • MiTec &gt; Graduación<br>
                        • MiTec &gt; Buscar "Servicio Social" &gt; Avance de horas
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      }
      break;
      
    case 'proposito_vida':
      if (estado === 'Sí') {
        // Versión COMPLETADA - Verde
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #E8F5E9; border: 1px solid #4CAF50;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">✅</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #2E7D32; font-size: 15px;">Propósito de vida actualizado</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → ¡Excelente! Tu propósito está actualizado en los últimos 6 meses.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      } else {
        // Versión PENDIENTE - Rojo
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 15px; background-color: #ffffff; border: 1px solid #dee2e6;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">❌</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #DC3545; font-size: 15px;">Propósito de vida actualizado</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → Tu última actualización fue hace más de 6 meses. Es momento de redefinir tu norte profesional.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      }
      break;
      
    case 'meta_ocupacional':
      if (estado === 'Sí') {
        // Versión COMPLETADA - Verde (sin margin-bottom porque es el último)
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #E8F5E9; border: 1px solid #4CAF50;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">✅</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #2E7D32; font-size: 15px;">Meta ocupacional reciente</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → ¡Perfecto! Tu meta ocupacional está actualizada en los últimos 6 meses.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      } else {
        // Versión PENDIENTE - Rojo
        html = `
          <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #ffffff; border: 1px solid #dee2e6;">
            <tr>
              <td style="padding: 12px;">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td style="width: 30px; vertical-align: top; text-align: center;">
                      <span style="font-size: 20px;">❌</span>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px;">
                      <strong style="color: #DC3545; font-size: 15px;">Meta ocupacional reciente</strong>
                      <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                        → Tu última actualización fue hace más de 6 meses. Actualiza tu plan de carrera para los próximos 2 años.
                      </p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>`;
      }
      break;
  }
  
  return html;
}

// Función principal para generar el correo completo versión 2 (con bordes redondeados, sin gradientes)
function generarCorreoPersonalizadoV2(estudiante, linkWebApp, nombreMentor) {
  // Extraer datos del estudiante
  const nombre = estudiante['Nombre'];
  const programa = estudiante['Programa'];
  const email = estudiante['Email'] || estudiante['Correo'];
  const ingles = estudiante['Inglés (Inv2025)'];
  const horasSS = estudiante['Horas de SS'] || 0;
  const propositoReciente = estudiante['Propósito de vida reciente'];
  const metaOcupacional = estudiante['Meta ocupacional reciente'];
  
  // Generar items del checklist
  const checklistIngles = generarItemChecklist('ingles', ingles, {});
  const checklistServicio = generarItemChecklist('servicio_social', '', {horas: horasSS});
  const checklistProposito = generarItemChecklist('proposito_vida', propositoReciente, {});
  const checklistMeta = generarItemChecklist('meta_ocupacional', metaOcupacional, {});
  
  // Contar pendientes para el asunto
  let pendientes = 0;
  if (!ingles || ingles === 'No cumple - SE' || ingles === 'No' || ingles === '-') pendientes++;
  if (horasSS < 480) pendientes++;
  if (propositoReciente !== 'Sí') pendientes++;
  if (metaOcupacional !== 'Sí') pendientes++;
  
  // HTML base del correo (copiar del artifact anterior pero con placeholders para los items dinámicos)
  let htmlCompleto = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Camino a tu Graduación - 12 Semanas</title>
    <!--[if mso]>
    <noscript>
        <xml>
            <o:OfficeDocumentSettings>
                <o:PixelsPerInch>96</o:PixelsPerInch>
            </o:OfficeDocumentSettings>
        </xml>
    </noscript>
    <![endif]-->
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5;">
    
    <!-- Contenedor Principal -->
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f5f5f5;">
        <tr>
            <td align="center" style="padding: 20px 0;">
                
                <!-- Contenedor del Email -->
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    
                    <!-- Header con color sólido azul TEC -->
                    <tr>
                        <td style="background-color: #002B7A; padding: 35px 40px; text-align: center;">
                            <h1 style="color: #ffffff; margin: 0 0 8px 0; font-size: 28px; font-weight: 600;">
                                🎓 Camino a tu Graduación
                            </h1>
                            <p style="color: #ffffff; margin: 0; font-size: 16px; opacity: 0.95;">
                                Tu checklist personalizado de requisitos
                            </p>
                        </td>
                    </tr>
                    
                    <!-- Banner de cuenta regresiva -->
                    <tr>
                        <td style="background-color: #E8F4FD; padding: 18px 40px; border-bottom: 2px solid #4A90E2;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td style="width: 45px; vertical-align: middle;">
                                        <span style="font-size: 32px;">⏰</span>
                                    </td>
                                    <td style="vertical-align: middle;">
                                        <strong style="color: #002B7A; font-size: 17px;">Faltan 12 semanas para tu ceremonia de graduación</strong>
                                        <p style="color: #54585A; margin: 3px 0 0 0; font-size: 14px;">Es momento de asegurar que todo esté en orden</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    
                    <!-- Saludo personalizado -->
                    <tr>
                        <td style="padding: 30px 40px 20px 40px;">
                            <h2 style="color: #002B7A; margin: 0 0 15px 0; font-size: 22px;">
                                Hola [Nombre],
                            </h2>
                            <p style="color: #54585A; margin: 0 0 15px 0; line-height: 1.6; font-size: 15px;">
                                ¡Felicidades por llegar al <strong>8vo semestre</strong> de <strong>[Programa]</strong>! 🎉
                            </p>
                            <p style="color: #54585A; margin: 0; line-height: 1.6; font-size: 15px;">
                                Hemos revisado tu expediente y preparado un checklist personalizado con los requisitos que necesitas completar para graduarte sin contratiempos este diciembre.
                            </p>
                        </td>
                    </tr>
                    
                    <!-- Checklist de Graduación DINÁMICO -->
                    <tr>
                        <td style="padding: 20px 40px;">
                            <div style="background-color: #F8F9FA; border-radius: 8px; padding: 25px; border-left: 4px solid #002B7A;">
                                <h3 style="color: #002B7A; margin: 0 0 20px 0; font-size: 18px;">
                                    📋 TU CHECKLIST PERSONALIZADO DE GRADUACIÓN:
                                </h3>
                                
                                <!-- Bloques dinámicos inyectados por generarCorreoPersonalizadoSinGradientes -->
                                {{CHECKLIST_INGLES}}
                                {{CHECKLIST_SERVICIO}}
                                {{CHECKLIST_PROPOSITO}}
                                {{CHECKLIST_META}}
                                
                                <!-- Item 4 ALTERNATIVO: Meta ocupacional (EJEMPLO SÍ CUMPLE) - Comentado
                                <div style="padding: 12px; background-color: #E8F5E9; border-radius: 6px; border: 1px solid #4CAF50;">
                                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                                        <tr>
                                            <td style="width: 30px; vertical-align: top;">
                                                <span style="font-size: 20px;">✅</span>
                                            </td>
                                            <td style="vertical-align: top;">
                                                <strong style="color: #2E7D32; font-size: 15px;">Meta ocupacional reciente</strong>
                                                <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">
                                                    → ¡Perfecto! Tu meta ocupacional está actualizada en los últimos 6 meses.
                                                </p>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                -->
                            </div>
                        </td>
                    </tr>
                    
                    <!-- Sección: Siguiente Paso -->
                    <tr>
                        <td style="padding: 25px 40px;">
                            <div style="background-color: #FFC107; border-radius: 8px; padding: 25px; text-align: center;">
                                <h3 style="color: #000000; margin: 0 0 12px 0; font-size: 20px;">
                                    🎯 Tu siguiente paso es muy sencillo
                                </h3>
                                <p style="color: #333333; margin: 0 0 18px 0; font-size: 15px; line-height: 1.5;">
                                    Queremos conocer cómo te sientes ante esta transición profesional.<br>
                                    <strong>Completa este breve test (solo 2 minutos)</strong> que nos ayudará a personalizar el apoyo que necesitas.
                                </p>
                                <a href="#" style="display: inline-block; padding: 14px 40px; background-color: #002B7A; color: #ffffff; text-decoration: none; border-radius: 25px; font-weight: 600; font-size: 16px; box-shadow: 0 3px 6px rgba(0,0,0,0.2);">
                                    INICIAR MI EVALUACIÓN →
                                </a>
                                <p style="color: #666666; margin: 15px 0 0 0; font-size: 13px;">
                                    Al completar el test, agendarás automáticamente tu sesión de mentoría del 17 de septiembre
                                </p>
                            </div>
                        </td>
                    </tr>
                    
                    <!-- Información de la sesión (más sutil) -->
                    <tr>
                        <td style="padding: 0 40px 25px 40px;">
                            <div style="background-color: #F0F7FF; border-radius: 8px; padding: 20px;">
                                <h4 style="color: #002B7A; margin: 0 0 12px 0; font-size: 16px;">
                                    ✨ ¿Qué incluye tu sesión personalizada?
                                </h4>
                                <p style="color: #54585A; margin: 0 0 12px 0; font-size: 14px; line-height: 1.6;">
                                    Tu mentor/a estudiantil te acompañará para:
                                </p>
                                <ul style="color: #54585A; margin: 0 0 15px 0; padding-left: 20px; font-size: 14px; line-height: 1.7;">
                                    <li>Revisar detalladamente cada requisito pendiente</li>
                                    <li>Crear un plan de acción con fechas específicas</li>
                                    <li>Trabajar en tu propósito de vida y meta ocupacional</li>
                                    <li>Prepararte para tu inserción al mundo laboral</li>
                                </ul>
                                
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="background-color: #ffffff; border-radius: 6px; padding: 12px; width: 100%;">
                                    <tr>
                                        <td style="font-size: 13px; color: #54585A; line-height: 1.5;">
                                            <strong>📍 Lugar:</strong> Mentoría Estudiantil, Centrales Sur, 3er Piso<br>
                                            <strong>📅 Fecha:</strong> 17 de septiembre de 2025<br>
                                            <strong>⏰ Duración:</strong> 45 minutos personalizados
                                        </td>
                                    </tr>
                                </table>
                            </div>
                        </td>
                    </tr>
                    
                    <!-- Mensaje motivacional -->
                    <tr>
                        <td style="padding: 0 40px 30px 40px;">
                            <div style="background-color: #FFFBF0; border-left: 3px solid #FFC107; padding: 15px 20px; border-radius: 4px;">
                                <p style="color: #54585A; margin: 0; font-size: 14px; line-height: 1.6;">
                                    <strong style="color: #002B7A;">💡 Recuerda:</strong> Tu inserción laboral exitosa depende de tener todos estos elementos en orden. Los empleadores valoran candidatos que demuestran preparación integral y claridad en sus objetivos profesionales.
                                </p>
                            </div>
                        </td>
                    </tr>
                    
                    <!-- Mensaje de cierre -->
                    <tr>
                        <td style="padding: 0 40px 35px 40px;">
                            <p style="color: #54585A; margin: 0 0 10px 0; line-height: 1.6; font-size: 15px;">
                                No estás solo/a en este proceso. Como tu mentor/a estudiantil, estoy aquí para asegurar que cruces la meta con éxito.
                            </p>
                            <p style="color: #54585A; margin: 15px 0 0 0; line-height: 1.6; font-size: 15px;">
                                <strong>Da el primer paso ahora:</strong> completa tu evaluación y nos vemos el 17 de septiembre.
                            </p>
                            <p style="color: #002B7A; margin: 25px 0 0 0; font-weight: 600; font-size: 15px;">
                                ¡Vamos por esa graduación!<br><br>
                                [Nombre del Mentor/a]
                            </p>
                            <p style="color: #6C757D; margin: 5px 0 0 0; font-size: 13px;">
                                Mentor/a Estudiantil<br>
                                Mentoría y Bienestar | Campus Monterrey
                            </p>
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #F8F9FA; padding: 20px 40px; text-align: center; border-top: 1px solid #DEE2E6;">
                            <p style="color: #6C757D; margin: 0 0 10px 0; font-size: 12px; font-style: italic;">
                                P.D. El 87% de los estudiantes que completan su evaluación y asisten a su sesión de mentoría logran graduarse sin contratiempos. ¡Sé parte de este grupo exitoso!
                            </p>
                            <p style="color: #ADB5BD; margin: 10px 0 0 0; font-size: 11px;">
                                Este correo fue enviado a [Email] como parte del programa de Mentoría y Bienestar<br>
                                © 2025 Tecnológico de Monterrey | Campus Monterrey
                            </p>
                        </td>
                    </tr>
                    
                </table>
                <!-- Fin del contenedor del email -->
                
            </td>
        </tr>
    </table>
    
</body>
</html>
  `;
  
  // Reemplazar variables
  htmlCompleto = htmlCompleto
    .replace(/\[Nombre\]/g, nombre)
    .replace(/\[Programa\]/g, programa)
    .replace(/\[Email\]/g, email)
    .replace(/\[Nombre del Mentor\/a\]/g, nombreMentor)
    .replace('{{CHECKLIST_INGLES}}', checklistIngles)
    .replace('{{CHECKLIST_SERVICIO}}', checklistServicio)
    .replace('{{CHECKLIST_PROPOSITO}}', checklistProposito)
    .replace('{{CHECKLIST_META}}', checklistMeta)
    .replace('href="#"', `href="${linkWebApp}"`);

  // Inyectar datos de sesión desde variables globales si existen
  try {
    var sesLoc = (typeof SESSION_LOCATION !== 'undefined' && SESSION_LOCATION) ? SESSION_LOCATION : null;
    var sesDateISO = (typeof SESSION_DATE !== 'undefined' && SESSION_DATE) ? SESSION_DATE : null;
    var sesStart = (typeof SESSION_TIME_START !== 'undefined' && SESSION_TIME_START) ? SESSION_TIME_START : null;
    var sesEnd = (typeof SESSION_TIME_END !== 'undefined' && SESSION_TIME_END) ? SESSION_TIME_END : null;
    var sesDateLarga = sesDateISO ? formatSpanishDateLong_(sesDateISO) : null;
    var sesDateCorta = sesDateISO ? formatSpanishDateShort_(sesDateISO) : null;
    var sesDurMin = (sesStart && sesEnd) ? minutesBetween_(sesStart, sesEnd) : null;

    if (sesLoc) {
      htmlCompleto = htmlCompleto.replace(/(<strong>📍 Lugar:<\/strong>) [^<]+/, `$1 ${sesLoc}`);
    }
    if (sesDateLarga) {
      htmlCompleto = htmlCompleto.replace(/(<strong>📅 Fecha:<\/strong>) [^<]+/, `$1 ${sesDateLarga}`);
      // Frases cortas en otros bloques
      htmlCompleto = htmlCompleto.replace(/del 17 de septiembre/g, `del ${sesDateCorta}`);
      htmlCompleto = htmlCompleto.replace(/nos vemos el 17 de septiembre\./g, `nos vemos el ${sesDateCorta}.`);
    }
    if (sesDurMin != null) {
      htmlCompleto = htmlCompleto.replace(/(<strong>⏰ Duración:<\/strong>) [^<]+/, `$1 ${sesDurMin} minutos personalizados`);
    }
  } catch (e) { /* best-effort */ }
  
  return {
    html: htmlCompleto,
    pendientes: pendientes
  };
}

// ======= Helpers de formato de fecha/tiempo =======
function minutesBetween_(hhmmStart, hhmmEnd) {
  function toMin(s){ var m=s.split(':'); return (+m[0])*60 + (+m[1]); }
  try { return Math.max(0, toMin(hhmmEnd) - toMin(hhmmStart)); } catch(_) { return 45; }
}

function formatSpanishDateLong_(iso) {
  var d = new Date(iso + 'T00:00:00');
  if (isNaN(d)) return iso;
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return d.getDate() + ' de ' + meses[d.getMonth()] + ' de ' + d.getFullYear();
}
function formatSpanishDateShort_(iso) {
  var d = new Date(iso + 'T00:00:00');
  if (isNaN(d)) return iso;
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return d.getDate() + ' de ' + meses[d.getMonth()];
}

// Ejemplo de uso
function testGenerarCorreo() {
  const estudianteEjemplo = {
    'Nombre': 'María',
    'Programa': 'LAE',
    'Email': 'maria@tec.mx',
    'Inglés (Inv2025)': 'Cumple (B2)',  // Este aparecerá en verde
    'Horas de SS': 480,  // Este aparecerá en verde
    'Propósito de vida reciente': 'Sí',  // Este aparecerá en verde
    'Meta ocupacional reciente': 'No'  // Este aparecerá en rojo
  };
  
  const resultado = generarCorreoPersonalizadoV2(
    estudianteEjemplo,
    'https://tu-webapp.com?id=A01234567',
    'Juan Pérez'
  );
  
  console.log(`Estudiante tiene ${resultado.pendientes} pendientes`);
  // El HTML generado mostrará 3 items en verde y 1 en rojo
}

// Alias para compatibilidad con Code.gs
function generarCorreoPersonalizadoSinGradientes(estudiante, linkWebApp, nombreMentor) {
  return generarCorreoPersonalizadoV2(estudiante, linkWebApp, nombreMentor);
}
