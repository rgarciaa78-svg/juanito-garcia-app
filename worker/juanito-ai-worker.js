// Cloudflare Worker — puente entre Juanito (juanito.html) y la API de Gemini.
// La API key vive en un "secret" del Worker (variable de entorno GEMINI_API_KEY),
// nunca en el HTML público, así nadie puede robarla ni usarla a tu costo.
//
// Para actualizarlo: Cloudflare → Workers & Pages → autumn-boat-225d →
// Edit code → pegar este archivo completo → Deploy.

const ALLOWED_ORIGIN = "https://rgarciaa78-svg.github.io";

// Modelos en orden de preferencia. El primero es el más capaz; si devuelve 503
// —"This model is currently experiencing high demand"— se prueba el siguiente.
// Sin esta lista, una saturación temporal de un modelo dejaba el chat entero
// sin IA y el usuario recibía una plantilla sin saber que el modelo nunca
// había contestado.
const MODELOS = [
  "gemini-3.6-flash",
  "gemini-2.5-flash",
  "gemini-2.0-flash",
];

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }
    if (request.method !== "POST") {
      return json({ error: "Método no permitido" }, 405);
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return json({ error: "JSON inválido" }, 400);
    }

    const { pregunta, contexto, historial } = body || {};
    if (!pregunta || !contexto) {
      return json({ error: "Falta 'pregunta' o 'contexto'" }, 400);
    }

    const systemPrompt = `Eres JUANITO, el agente de Business Intelligence del holding del usuario
(empresas: PAUNO, AMAUTA, NEOPACK). Respondes preguntas de gerencia/CEO sobre KPIs, mermas, mora,
compras, margen, inventario, control interno, etc.

REGLAS INNEGOCIABLES:
1. Usa SOLO los datos del JSON de contexto. Nunca inventes cifras.
2. Si el dato no está, dilo en una frase y ofrece la pregunta más cercana que sí puedas responder.
3. NUNCA completes un dato faltante con promedios de industria, benchmarks o supuestos.
4. El contexto trae un campo _guia con instrucciones de cómo razonar: síguelas.
5. El contexto trae series mensuales en 'series': úsalas para responder "desde cuándo"
   y "cómo viene", y los desgloses para responder "de dónde viene".
6. Cita el mes de cada cifra. Los períodos están en series.periodos, alineados por
   posición con cada array de valores.
7. Revisa 'avisos_datos' antes de afirmar que algo es actual.

Responde en español, tono ejecutivo y directo, con números concretos citados del contexto.
Usa HTML simple para formato (negritas <strong>, saltos <br>, listas simples) — no markdown.`;

    const contents = [];
    (historial || []).slice(-6).forEach(m => {
      contents.push({
        role: m.rol === "assistant" ? "model" : "user",
        parts: [{ text: m.texto }]
      });
    });
    contents.push({
      role: "user",
      parts: [{ text: `CONTEXTO (JSON con los datos reales del dashboard):\n${contexto}\n\nPREGUNTA: ${pregunta}` }]
    });

    const cuerpo = JSON.stringify({
      system_instruction: { parts: [{ text: systemPrompt }] },
      contents,
      generationConfig: {
        temperature: 0.3,
        // Los modelos con razonamiento gastan parte de este presupuesto
        // pensando ANTES de escribir, y ese consumo no se ve en la respuesta.
        // Con 2048 las contestaciones salían cortadas a media frase; con 8192
        // seguían cortándose, así que el razonamiento se llevaba casi todo.
        maxOutputTokens: 16384,
        // Se acota lo que puede pensar para que quede presupuesto real para
        // escribir. Estas preguntas se responden leyendo series y desgloses:
        // no necesitan cadenas de razonamiento largas.
        thinkingConfig: { thinkingBudget: 1024 }
      }
    });

    let ultimoError = "";
    for (const modelo of MODELOS) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent`;
      let resp;
      try {
        resp = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
          body: cuerpo
        });
      } catch (e) {
        ultimoError = `fallo de red con ${modelo}: ${e}`;
        continue;
      }

      if (resp.ok) {
        const data = await resp.json();
        const cand = data?.candidates?.[0];
        let texto = cand?.content?.parts?.map(p => p.text).join("") || "";
        if (!texto) {
          ultimoError = `${modelo} devolvió una respuesta vacía `
            + `(finishReason: ${cand?.finishReason || "desconocido"})`;
          continue;
        }
        // Si el modelo se quedó sin presupuesto, la respuesta llega cortada a
        // media frase. Es mejor decirlo que dejar al lector con media idea.
        if (cand?.finishReason === "MAX_TOKENS") {
          texto += "\n\n[Respuesta cortada por límite de longitud. "
                 + "Pregunta por una parte concreta para verla completa.]";
        }
        // Por si responde en markdown pese a la instrucción.
        texto = texto.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>").replace(/\n/g, "<br>");
        // finishReason y el conteo de tokens se devuelven siempre: sin ellos,
        // diagnosticar por qué una respuesta sale corta es adivinar.
        return json({
          respuesta: texto,
          modelo,
          finishReason: cand?.finishReason || null,
          tokens: data?.usageMetadata || null
        });
      }

      ultimoError = await resp.text();
      // Saturación o límite de cuota: tiene sentido probar otro modelo.
      // Cualquier otro error (clave inválida, petición mal formada) se
      // repetiría igual en todos, así que se corta aquí.
      const reintentable = resp.status === 503 || resp.status === 429 ||
                           /UNAVAILABLE|high demand|overload/i.test(ultimoError);
      if (!reintentable) break;
    }

    return json({ error: "Error de Gemini", detalle: ultimoError.slice(0, 600) }, 502);
  }
};

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() }
  });
}
