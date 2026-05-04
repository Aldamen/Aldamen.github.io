const projectsData = [
    // PROYECTO 1: Empleo anterior (Abstraído)
    /*"Desarrollo de un pipeline de análisis de datos para optimizar la proyección del Unit Cash Cost (Costo Unitario en Efectivo). Procesé grandes volúmenes de datos históricos de operación de equipos pesados, desarrollando scripts en Python para ajustar automáticamente distribuciones de probabilidad a las variables operativas.<br><br>Los parámetros estadísticos obtenidos se integraron en modelos financieros en Excel, permitiendo ejecutar simulaciones de Monte Carlo. Esto transformó proyecciones estáticas en herramientas dinámicas de análisis de riesgo, brindando a la gerencia un soporte cuantitativo robusto para la toma de decisiones bajo incertidumbre."*/
    {
        id: "p1",
        es: {
            title: "Modelo Probabilístico y Análisis de Riesgo",
            shortDesc: "Simulación de Costos con Python y Excel",
            company: "International Project Management (Cerrejón)",
            sector: "Planeación Estratégica / Modelado y Simulación",
            tools: "Python, Pandas, Excel, SciPy, SQL, @RISK",
            outcome: "+15% Precisión en Pronósticos",
            body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                En la industria minera, el <strong>Unit Cash Cost (UCC)</strong> es una métrica muy importante pero difícil de predecir, ya que depende del rendimiento de los equipos, el precio de los insumos y las condiciones del terreno. En este proyecto, me propuse resolver un problema clave: la productividad que planeábamos para los ciclos de carga y transporte (pala-camión) casi nunca se cumplía en la operación real, lo que dificultaba mucho la toma de decisiones.<br><br>
                
                Para dejar de tomar decisiones basadas en suposiciones, me enfoqué en los datos. Extraje más de <strong>7 años de registros históricos</strong> desde una base de datos <strong>SQL</strong>. Mi objetivo fue analizar toda esta información para descubrir exactamente qué factores estaban afectando nuestra productividad y encontrar formas de solucionarlo directamente en el campo.<br><br>
                
                Este proyecto nos permitió evolucionar la forma de trabajar en el área: dejamos atrás los cálculos estáticos (que nos daban un solo número fijo) y pasamos a realizar proyecciones más reales basadas en un rango de probabilidades.<br><br>
                
                <h4>Lo que hice</h4>
                Desarrollé una metodología que combina la capacidad de procesamiento de Python con los modelos financieros de Excel:
                <ul>
                    <li><strong>Análisis de Datos:</strong> Utilicé Python y Pandas para procesar los años de historia operativa extraídos de SQL, con el fin de identificar tendencias reales de rendimiento de los equipos.</li>
                    <li><strong>Ajuste Estadístico:</strong> A través de SciPy, automaticé la búsqueda del modelo matemático que mejor representaba el comportamiento de cada variable de la mina.</li>
                    <li><strong>Simulación de Escenarios:</strong> Integré estos resultados en Excel para ejecutar miles de simulaciones (Monte Carlo), lo que me permitió predecir cómo se comportaría el costo bajo diferentes situaciones.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Limpieza de Datos (Pandas)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Este script fue diseñado para filtrar automáticamente los datos atípicos o errores de registro de los equipos. Esto es esencial para asegurar que los cálculos no se vean afectados por números que son imposibles en la operación real.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            limpieza_outliers.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>import pandas as pd

combinations = [
    ("Pala1", "Camion1"), ("Pala1", "Camion2"), ("Pala1", "Camion3"),
    ("Pala2", "Camion1"), ("Pala2", "Camion2"), ("Pala2", "Camion3"),
    ("Pala3", "Camion1"), ("Pala3", "Camion2"), ("Pala3", "Camion3"),
]

file_path = "Palas_y_camiones.xlsx"  
data = pd.ExcelFile(file_path)
df = data.parse("Datos")

columns_to_filter = ["Cargas", "Toneladas", "Tiempo1", "Tiempo2",
                     "Tiempo3", "Tiempo4", "FC", "Productividad"]

def filter_outliers(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]

filtered_data = {}

for pala, camion in combinations:
    subset = df[(df["Flota de Pala"] == pala) & (df["Flota de Camion"] == camion)]
    for column in columns_to_filter:
        if column in subset.columns:
            subset = filter_outliers(subset, column)
    filtered_data[f"{pala}_{camion}"] = subset

output_path = "Palas_y_camiones_filtrados.xlsx"
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, data_frame in filtered_data.items():
        data_frame.to_excel(writer, index=False, sheet_name=sheet_name)</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por acuerdos de confidencialidad, este código es solo un fragmento importante del desarrollo real y los nombres han sido modificados.
                    </div>
                </div>

                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Simulador de Distribución</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Esta herramienta visual permite mostrar de forma interactiva lo que el código en Python realiza en segundo plano: probar distintos parámetros matemáticos hasta encontrar la curva que mejor encaje con la realidad de los datos operativos.</p>
                    
                    <div id="sim-wrapper" style="background: #111; padding: 25px; border-radius: 12px; border: 1px solid #333; width: 100%; box-sizing: border-box; font-family: system-ui, -apple-system, sans-serif;">
                        <h3 style="text-align: center; color: #fff; margin: 0 0 20px 0; font-weight: 500;">Ajuste de Distribución de Equipos</h3>
                        
                        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px;">
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Media</small>
                                <span id="txt-media" style="font-size: 1.4rem; color: #fff; font-weight: bold;">10.0</span>
                            </div>
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Volatilidad</small>
                                <span id="txt-std" style="font-size: 1.4rem; color: #fff; font-weight: bold;">2.0</span>
                            </div>
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Precisión del Ajuste</small>
                                <span id="txt-acc" style="font-size: 1.4rem; color: #ff3333; font-weight: bold;">0%</span>
                            </div>
                        </div>

                        <div style="height: 330px; width: 100%; margin-bottom: 25px;">
                            <canvas id="simChart"></canvas>
                        </div>

                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 20px;">
                            <div>
                                <label style="color: #ccc; font-size: 0.9rem; display: block; margin-bottom: 10px;">Media (Rendimiento)</label>
                                <input type="range" id="input-mean" min="0" max="100" value="10" style="width: 100%; accent-color: #1e40af;">
                            </div>
                            <div>
                                <label style="color: #ccc; font-size: 0.9rem; display: block; margin-bottom: 10px;">Desviación (Volatilidad)</label>
                                <input type="range" id="input-std" min="1" max="20" step="0.5" value="2" style="width: 100%; accent-color: #1e40af;">
                            </div>
                        </div>

                        <div id="sim-msg" style="padding: 15px; border-radius: 8px; background: rgba(255,51,51,0.1); color: #ff3333; border: 1px solid rgba(255,51,51,0.2); font-size: 0.95rem; text-align: center;">
                            Ajusta los parámetros para iniciar la validación operativa.
                        </div>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por políticas de privacidad, este simulador es únicamente una representación visual de la herramienta y no contiene datos reales de la operación.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                <h4>Resultados del Proyecto</h4>
                Con la implementación de este modelo logré aportar soluciones claras para la operación diaria de la mina:
                <ul>
                    <li><strong>Identificación de Problemas:</strong> Encontré cuáles eran las variables que realmente estaban frenando la operación, lo que permitió al equipo enfocar sus esfuerzos de mejora en el lugar correcto.</li>
                    <li><strong>Mayor Precisión:</strong> Logré aumentar la precisión de los pronósticos de costos en un 15% al compararlo con cómo se hacía antes.</li>
                    <li><strong>Mejor Planificación:</strong> Entregué gráficos claros que permitían a la gerencia ver fácilmente qué tan probable era cumplir con el presupuesto anual o las metas de extracción.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">3. Selección del Mejor Modelo (SciPy)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">El cálculo del AIC (Criterio de Información de Akaike) es una pieza clave del desarrollo. Es la parte del algoritmo responsable de comparar matemáticamente varias distribuciones y escoger objetivamente la que mejor predice el futuro, sin depender de opiniones.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            analisis_distribuciones.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.8rem; line-height: 1.4;"><code>import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
from scipy.stats import logistic

def calcular_aic(columna, distribucion, parametros):
    n = len(columna)
    k = len(parametros)
    log_likelihood = np.sum(distribucion.logpdf(columna, *parametros))
    aic = 2 * k - 2 * log_likelihood
    return round(aic, 5)

def realizar_pruebas_de_ajuste(archivo_excel, nombre_hoja, num_cols, ini, fin):
    data = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
    data = data.iloc[ini-1:fin]

    distribuciones = {
        "Normal": stats.norm,
        "Lognormal": stats.lognorm,
        "Logistic": logistic
    }

    for columna_num in num_cols:
        columna = data.iloc[:, columna_num]
        columna = columna[np.isfinite(columna)]

        for nombre, dist en distribuciones.items():
            parametros = dist.fit(columna)
            x = np.linspace(columna.min(), columna.max(), 100)
            pdf = dist.pdf(x, *parametros)

            fig, ax = plt.subplots()
            ax.hist(columna, bins=20, density=True, color='g')
            ax.plot(x, pdf, label=f'Ajuste {nombre}')
            plt.show()</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por acuerdos de confidencialidad, este código es solo un fragmento importante del desarrollo real y los nombres han sido modificados.
                    </div>
                </div>

                <div>
    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">4. Variables que más afectan (Tornado)</h5>
    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Este gráfico muestra de forma sencilla qué factores afectan más a la productividad y a los volúmenes de extracción. Resulta muy útil para comparar el rendimiento de periodos recientes contra todo el histórico de 7 años almacenado en la base de datos.</p>
    
    <div style="background: #111; border: 1px solid #333; border-radius: 12px; display: flex; align-items: center; justify-content: center; padding: 15px; height: calc(100% - 90px);">
        <img src="assets/images/tornado.jpg" alt="Gráfico de Tornado de Sensibilidades" class="clickable-img" style="width: 100%; max-height: 450px; object-fit: contain; border-radius: 8px;">
    </div>
    
    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
        * Nota: Por políticas de privacidad, esta imagen es una representación visual y no expone sensibilidades reales del negocio.
    </div>
</div>

            </div>`
        },
        en: {
            title: "Probabilistic Model & Risk Analysis",
            shortDesc: "Cost Simulation with Python and Excel",
            company: "International Project Management (Cerrejón)",
            sector: "Strategic Planning / Modeling and Simulation",
            tools: "Python, Pandas, Excel, SciPy, SQL, @RISK",
            outcome: "+15% Forecast Accuracy",
            body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                In the mining industry, <strong>Unit Cash Cost (UCC)</strong> is a critical but hard-to-predict metric because it depends on equipment performance, material prices, and terrain conditions. In this project, I set out to solve a key problem: the productivity we planned for shovel-truck cycles was rarely met in real operations, which made decision-making very difficult.<br><br>
                
                To stop making decisions based on assumptions, I focused on the data. I extracted over <strong>7 years of historical records</strong> from a <strong>SQL</strong> database. My goal was to analyze all this information to discover exactly what factors were affecting our productivity and find ways to fix it directly in the field.<br><br>
                
                This project helped us evolve the way we worked: we left behind static calculations (which gave us a single fixed number) and moved to more realistic projections based on a range of probabilities.<br><br>
                
                <h4>What I Did</h4>
                I developed a methodology that combines Python's processing capabilities with Excel's financial models:
                <ul>
                    <li><strong>Data Analysis:</strong> I used Python and Pandas to process the years of operational history extracted from SQL, aiming to identify real equipment performance trends.</li>
                    <li><strong>Statistical Fitting:</strong> Using SciPy, I automated the search for the mathematical model that best represented the behavior of each variable in the mine.</li>
                    <li><strong>Scenario Simulation:</strong> I integrated these results into Excel to run thousands of simulations (Monte Carlo), allowing me to predict how the cost would behave in different situations.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Data Cleaning (Pandas)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This script was designed to automatically filter out outliers or equipment recording errors. This is essential to ensure that calculations are not skewed by numbers that are impossible in real-world operations.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            outlier_cleaning.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>import pandas as pd

combinations = [
    ("Pala1", "Camion1"), ("Pala1", "Camion2"), ("Pala1", "Camion3"),
    ("Pala2", "Camion1"), ("Pala2", "Camion2"), ("Pala2", "Camion3"),
    ("Pala3", "Camion1"), ("Pala3", "Camion2"), ("Pala3", "Camion3"),
]

file_path = "Palas_y_camiones.xlsx"  
data = pd.ExcelFile(file_path)
df = data.parse("Datos")

columns_to_filter = ["Cargas", "Toneladas", "Tiempo1", "Tiempo2",
                     "Tiempo3", "Tiempo4", "FC", "Productividad"]

def filter_outliers(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]

filtered_data = {}

for pala, camion in combinations:
    subset = df[(df["Flota de Pala"] == pala) & (df["Flota de Camion"] == camion)]
    for column in columns_to_filter:
        if column in subset.columns:
            subset = filter_outliers(subset, column)
    filtered_data[f"{pala}_{camion}"] = subset

output_path = "Palas_y_camiones_filtrados.xlsx"
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, data_frame in filtered_data.items():
        data_frame.to_excel(writer, index=False, sheet_name=sheet_name)</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to confidentiality agreements, this code is only a representative fragment and names have been altered.
                    </div>
                </div>

                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Distribution Simulator</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This visual tool interactively shows what the Python code does behind the scenes: testing different mathematical parameters until finding the curve that best fits the reality of the operational data.</p>
                    
                    <div id="sim-wrapper" style="background: #111; padding: 25px; border-radius: 12px; border: 1px solid #333; width: 100%; box-sizing: border-box; font-family: system-ui, -apple-system, sans-serif;">
                        <h3 style="text-align: center; color: #fff; margin: 0 0 20px 0; font-weight: 500;">Equipment Distribution Fitting</h3>
                        
                        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px;">
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Mean</small>
                                <span id="txt-media" style="font-size: 1.4rem; color: #fff; font-weight: bold;">10.0</span>
                            </div>
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Volatility</small>
                                <span id="txt-std" style="font-size: 1.4rem; color: #fff; font-weight: bold;">2.0</span>
                            </div>
                            <div style="text-align: center; background: #1a1a1a; padding: 10px; border-radius: 8px;">
                                <small style="color: #888; display: block; margin-bottom: 5px;">Fitting Accuracy</small>
                                <span id="txt-acc" style="font-size: 1.4rem; color: #ff3333; font-weight: bold;">0%</span>
                            </div>
                        </div>

                        <div style="height: 330px; width: 100%; margin-bottom: 25px;">
                            <canvas id="simChart"></canvas>
                        </div>

                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 20px;">
                            <div>
                                <label style="color: #ccc; font-size: 0.9rem; display: block; margin-bottom: 10px;">Mean (Performance)</label>
                                <input type="range" id="input-mean" min="0" max="100" value="10" style="width: 100%; accent-color: #1e40af;">
                            </div>
                            <div>
                                <label style="color: #ccc; font-size: 0.9rem; display: block; margin-bottom: 10px;">Deviation (Volatility)</label>
                                <input type="range" id="input-std" min="1" max="20" step="0.5" value="2" style="width: 100%; accent-color: #1e40af;">
                            </div>
                        </div>

                        <div id="sim-msg" style="padding: 15px; border-radius: 8px; background: rgba(255,51,51,0.1); color: #ff3333; border: 1px solid rgba(255,51,51,0.2); font-size: 0.95rem; text-align: center;">
                            Adjust the parameters to start the operational validation.
                        </div>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to privacy policies, this simulator is solely a visual representation of the tool and does not contain real operational data.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                <h4>Project Results</h4>
                With the implementation of this model, I was able to provide clear solutions for the mine's daily operations:
                <ul>
                    <li><strong>Problem Identification:</strong> I found which variables were actually slowing down the operation, allowing the team to focus their improvement efforts in the right place.</li>
                    <li><strong>Higher Accuracy:</strong> I increased cost forecasting accuracy by 15% compared to how it was done previously.</li>
                    <li><strong>Better Planning:</strong> I provided clear charts that allowed management to easily see how likely it was to meet the annual budget or extraction goals.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">3. Selecting the Best Model (SciPy)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Calculating the AIC (Akaike Information Criterion) is a key piece of the development. It is the part of the code responsible for mathematically comparing various distributions and objectively choosing the one that best predicts the future, without relying on opinions.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            distribution_analysis.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.8rem; line-height: 1.4;"><code>import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
from scipy.stats import logistic

def calcular_aic(columna, distribucion, parametros):
    n = len(columna)
    k = len(parametros)
    log_likelihood = np.sum(distribucion.logpdf(columna, *parametros))
    aic = 2 * k - 2 * log_likelihood
    return round(aic, 5)

def realizar_pruebas_de_ajuste(archivo_excel, nombre_hoja, num_cols, ini, fin):
    data = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
    data = data.iloc[ini-1:fin]

    distribuciones = {
        "Normal": stats.norm,
        "Lognormal": stats.lognorm,
        "Logistic": logistic
    }

    for columna_num in num_cols:
        columna = data.iloc[:, columna_num]
        columna = columna[np.isfinite(columna)]

        for nombre, dist en distribuciones.items():
            parametros = dist.fit(columna)
            x = np.linspace(columna.min(), columna.max(), 100)
            pdf = dist.pdf(x, *parametros)

            fig, ax = plt.subplots()
            ax.hist(columna, bins=20, density=True, color='g')
            ax.plot(x, pdf, label=f'Ajuste {nombre}')
            plt.show()</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to confidentiality agreements, this code is only a representative fragment and names have been altered.
                    </div>
                </div>

                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">4. Top Impacting Variables (Tornado)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This chart simply shows which factors have the biggest impact on productivity and extraction volumes. It was very useful for comparing performance from recent periods against the entire 7-year history stored in the database.</p>
                    
                    <div style="background: #111; border: 1px solid #333; border-radius: 12px; display: flex; align-items: center; justify-content: center; padding: 15px; height: calc(100% - 90px);">
                        <img src="assets/images/tornado.jpg" alt="Sensitivity Tornado Chart" class="clickable-img" style="width: 100%; max-height: 450px; object-fit: contain; border-radius: 8px;">
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to privacy policies, this image is a visual representation and does not expose real business sensitivities.
                    </div>
                </div>

            </div>`
        },
        gradient: "url('assets/images/proyecto1.png')"
    },

    // PROYECTO 2: Empleo anterior (Abstraído)
    {
    id: "p2",
    es: {
        title: "Automatización ETL y Gestión de Compras",
        shortDesc: "Dashboard Dinámico para Seguimiento de Especificaciones Técnicas",
        company: "International Project Management (Cerrejón)",
        sector: "Cadena de Suministro / Adquisición de Activos",
        tools: "SQL Server, Python, Power BI, Excel",
        outcome: "Reducción de Tiempos de Aprobación",
        body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                El proceso de compra de equipos mineros de gran escala (como camiones, palas y tractores) requiere de mucha coordinación. Antes de realizar cualquier adquisición, se crea un documento llamado Especificación Técnica (ET). Este documento debe ser revisado, validado y aprobado por especialistas de diversas áreas, como Seguridad, Activos y Mantenimiento, para asegurar que el equipo cumpla con todos los estándares operativos de la mina.<br><br>

                El problema principal era que el seguimiento de este proceso se hacía de forma manual. Una vez que enviábamos el documento a las diferentes áreas, perdíamos su rastro. No sabíamos con certeza quién ya lo había revisado, quién faltaba por aprobarlo, o qué persona estaba retrasando el trámite. Esta falta de control causaba demoras importantes en la compra de maquinaria que era vital para el trabajo diario de la mina.<br><br>

                Para solucionar esto, diseñé un sistema automatizado conectándome directamente a las bases de datos de la empresa. Programé un código en Python que se encargaba de extraer y actualizar la información todos los días a las 8:00 AM. Así, el equipo de compras ya no tenía que buscar y rastrear los documentos uno por uno, sino que al llegar a la oficina tenían un tablero interactivo con los datos actualizados listos para tomar decisiones rápidas.<br><br>

                <h4>Lo que hice</h4>
                <ul>
                    <li><strong>Extracción de Datos:</strong> Escribí consultas en SQL Server para cruzar la información de los documentos con los empleados asignados y calcular automáticamente cuántos días llevaba detenido cada proceso.</li>
                    <li><strong>Automatización Diaria:</strong> Desarrollé un script de Python que se ejecutaba como una tarea programada, refrescando toda la base de datos y el reporte todos los días por la mañana.</li>
                    <li><strong>Visualización de Gestión:</strong> Diseñé un dashboard interactivo que clasificaba el estado de cada documento y mostraba visualmente a las personas que tenían tareas pendientes, reduciendo los tiempos de espera.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Extracción y Cálculo de Tiempos (SQL Server)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Esta consulta extrae el historial de los documentos y calcula los días que han pasado desde que se asignó la tarea a una persona. Es la base de datos limpia que alimenta todo el sistema.</p>

                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            extraccion_et_flujo.sql
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>SELECT 
    et.ID_Especificacion,
    emp.Nombre_Completo AS Responsable,
    area.Nombre_Area AS Area,
    et.Codigo_ET + ' - ' + et.Descripcion AS Especificacion_Tecnica,
    flujo.Fase_Actual AS Estado_Fase,
    flujo.Estatus AS Condicion,
    DATEDIFF(day, flujo.Fecha_Asignacion, GETDATE()) AS Dias_Transcurridos
FROM 
    Tbl_Especificaciones et
INNER JOIN 
    Tbl_Flujo_Aprobacion flujo ON et.ID_Especificacion = flujo.ID_Especificacion
INNER JOIN 
    Tbl_Empleados emp ON flujo.Responsable_ID = emp.ID_Empleado
INNER JOIN 
    Tbl_Areas area ON emp.ID_Area = area.ID_Area
WHERE 
    flujo.Estatus = 'Pendiente'
ORDER BY 
    Dias_Transcurridos DESC;</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por acuerdos de confidencialidad, este código es una representación simplificada y los nombres de tablas y columnas han sido modificados.
                    </div>
                </div>

                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Dashboard de Monitoreo Diario</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Este es el panel que se actualizaba solo todas las mañanas. Agrupa las tareas pendientes en Validaciones, Endosos y Aprobaciones, mostrando claramente qué persona está retrasando el proceso de compra.</p>

                    <div style="background: #111; border: 1px solid #333; border-radius: 12px; display: flex; align-items: center; justify-content: center; padding: 15px; height: calc(100% - 90px);">
                        <img src="assets/images/Dashboard.jpg" alt="Dashboard de Seguimiento de Especificaciones Técnicas" class="clickable-img" style="width: 100%; max-height: 450px; object-fit: contain; border-radius: 8px;">
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por políticas de privacidad, los datos mostrados en este dashboard son ficticios y sirven únicamente como representación visual del producto final.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -20px;">
                <h4>Resultados del Proyecto</h4>
                Esta automatización cambió por completo la forma en que se gestionaban las adquisiciones en el área:
                <ul>
                    <li><strong>Visibilidad Total:</strong> El equipo ahora identifica al instante en qué escritorio está detenido cada documento técnico.</li>
                    <li><strong>Reducción de Tiempos:</strong> Al tener un seguimiento diario automático, los tiempos de espera entre fases se redujeron drásticamente al agilizar la comunicación con los responsables.</li>
                    <li><strong>Gestión Proactiva:</strong> El área de compras pasó de esperar respuestas a liderar el flujo de aprobación, acelerando la llegada de los equipos a la mina.</li>
                </ul>
            </div>`
    },
    en: {
        title: "ETL Automation & Procurement Management",
        shortDesc: "Dynamic Dashboard for Technical Specification Tracking",
        company: "International Project Management (Cerrejón)",
        sector: "Supply Chain / Asset Acquisition",
        tools: "SQL Server, Python, Power BI, Excel",
        outcome: "Reduction in Approval Times",
        body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                The process of purchasing large-scale mining equipment (like trucks, shovels, and tractors) requires careful coordination. Before any purchase is made, a Technical Specification (ET) document is created. This document must be reviewed, validated, and approved by specialists from various departments, such as Safety, Assets, and Maintenance, to ensure the equipment meets all operational standards of the mine.<br><br>

                The main problem was that tracking this process was done manually. Once we sent the document to the different departments, we lost track of its progress. We didn't know for sure who had already reviewed it, who still needed to approve it, or who was delaying the procedure. This lack of control caused significant delays in buying machinery that was vital for daily operations.<br><br>

                To solve this, I designed an automated system connecting directly to the company's databases. I programmed a Python script that extracted and updated the information every day at 8:00 AM. This way, the procurement team no longer had to track documents one by one. Instead, when they arrived at the office, they had an interactive dashboard with updated data ready for quick decision-making.<br><br>

                <h4>What I Did</h4>
                <ul>
                    <li><strong>Data Extraction:</strong> I wrote SQL Server queries to cross-reference document information with assigned employees and automatically calculate how many days each process had been delayed.</li>
                    <li><strong>Daily Automation:</strong> I developed a Python script that ran as a scheduled task, refreshing the entire database and the report every morning.</li>
                    <li><strong>Management Visualization:</strong> I designed an interactive dashboard that classified the status of each document and visually highlighted the people who had pending tasks, significantly reducing wait times.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Data Extraction & Time Calculation (SQL Server)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This query extracts document history and calculates the days passed since the task was assigned to a person. It is the clean database that feeds the entire system.</p>

                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            extract_workflow_et.sql
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>SELECT 
    et.Specification_ID,
    emp.Full_Name AS Responsible,
    area.Area_Name AS Area,
    et.ET_Code + ' - ' + et.Description AS Technical_Specification,
    workflow.Current_Phase AS Phase_Status,
    workflow.Condition AS Status,
    DATEDIFF(day, workflow.Assignment_Date, GETDATE()) AS Days_Elapsed
FROM 
    Tbl_Specifications et
INNER JOIN 
    Tbl_Approval_Workflow workflow ON et.Specification_ID = workflow.Specification_ID
INNER JOIN 
    Tbl_Employees emp ON workflow.Responsible_ID = emp.Employee_ID
INNER JOIN 
    Tbl_Areas area ON emp.Area_ID = area.Area_ID
WHERE 
    workflow.Condition = 'Pending'
ORDER BY 
    Days_Elapsed DESC;</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to confidentiality agreements, this code is a simplified representation and the names of tables and columns have been altered.
                    </div>
                </div>

                <div>
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Daily Monitoring Dashboard</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This is the panel that updated itself every morning. It groups pending tasks into Validations, Endorsements, and Approvals, clearly showing which person is delaying the purchasing process.</p>

                    <div style="background: #111; border: 1px solid #333; border-radius: 12px; display: flex; align-items: center; justify-content: center; padding: 15px; height: calc(100% - 90px);">
                        <img src="assets/images/Dashboard.jpg" alt="Technical Specification Tracking Dashboard" class="clickable-img" style="width: 100%; max-height: 450px; object-fit: contain; border-radius: 8px;">
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to privacy policies, the data shown in this dashboard is fictitious and serves solely as a visual representation of the final product.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -20px;">
                <h4>Project Results</h4>
                This automation completely changed the way procurement was managed in the department:
                <ul>
                    <li><strong>Total Visibility:</strong> The team can now instantly identify on whose desk each technical document is stalled.</li>
                    <li><strong>Time Reduction:</strong> With daily and automated tracking, wait times between phases were drastically reduced by streamlining communication with stakeholders.</li>
                    <li><strong>Proactive Management:</strong> The procurement team shifted from waiting for responses to leading the approval flow, speeding up the arrival of critical equipment at the mine.</li>
                </ul>
            </div>`
    },
    gradient: "url('assets/images/proyecto2.png')"
},

    // PROYECTO 3: Empleo anterior (Abstraído)
    {
    id: "p3",
    es: {
        title: "Value Driver Tree y Simulación de EBITDA",
        shortDesc: "Modelo Estocástico para Identificación de Fugas de Capital",
        company: "International Project Management (Cerrejón)",
        sector: "Planeación Financiera / Control de Gestión",
        tools: "Excel, @Risk, Python, Modelado Financiero",
        outcome: "Identificación del Driver Principal de Gasto",
        body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                En la gestión financiera de una mina, entender qué factor operativo afecta más nuestras ganancias (el EBITDA) es un gran desafío, ya que todo está conectado. En este proyecto me encargué de construir un informe de gerencia basado en un modelo tipo árbol (Value Driver Tree). La idea de este árbol es que comienza con el resultado financiero final y se va abriendo en ramas cada vez más pequeñas, hasta llegar a variables operativas muy específicas, como los tiempos en segundos que tarda un equipo en hacer su tarea.<br><br>

                El reto era que usar promedios fijos para planear no reflejaba la realidad de la mina. Para solucionar esto, utilicé la herramienta @Risk. Lo que hice fue tomar cada variable operativa del árbol (como los tiempos en segundos reales) y ponerla a variar entre un máximo y un mínimo. Luego, el modelo calculaba automáticamente cómo esa variabilidad alteraba los resultados de todas las demás ramas hacia arriba, hasta llegar al EBITDA de toda la organización.<br><br>

                De esta manera, dejamos de adivinar qué parte de la operación nos costaba más dinero y pasamos a tener un modelo matemático que nos decía exactamente dónde estábamos perdiendo rentabilidad.<br><br>

                <h4>Lo que hice</h4>
                <ul>
                    <li><strong>Modelado tipo Árbol:</strong> Estructuré un modelo financiero que conectaba las variables operativas en terreno (tiempos de ciclo, demoras en segundos) con los estados de pérdidas y ganancias de la mina.</li>
                    <li><strong>Simulación de Variabilidad (@Risk):</strong> Asigné rangos reales (mínimos y máximos en segundos) a las variables base y corrí miles de iteraciones para simular el comportamiento diario de la operación.</li>
                    <li><strong>Análisis de Sensibilidad:</strong> Extraje los resultados en un gráfico de tornado interactivo que le mostró a la gerencia cuál era la variable exacta que más estaba afectando el gasto.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Modelo de Simulación Estocástica</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Este script replica la lógica que utilicé para inyectar variabilidad entre máximos y mínimos (en segundos) en las ramas base del árbol, calculando el impacto final en el EBITDA a través de miles de iteraciones.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            vdt_montecarlo_ebitda.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>import numpy as np
import pandas as pd

variables_operativas = {
    'tiempo_cargue_seg': (120, 145, 180),
    'tiempo_acarreo_seg': (600, 650, 800),
    'demoras_operativas_seg': (45, 60, 120)
}

iteraciones = 10000
resultados_ebitda = []

for i in range(iteraciones):
    t_cargue = np.random.triangular(*variables_operativas['tiempo_cargue_seg'])
    t_acarreo = np.random.triangular(*variables_operativas['tiempo_acarreo_seg'])
    demoras = np.random.triangular(*variables_operativas['demoras_operativas_seg'])
    
    tiempo_ciclo_total = t_cargue + t_acarreo + demoras
    viajes_por_hora = 3600 / tiempo_ciclo_total
    
    toneladas_por_viaje = 320
    precio_tonelada = 115
    costo_operativo_hora = 450
    
    ingresos = viajes_por_hora * toneladas_por_viaje * precio_tonelada
    ebitda_simulado = ingresos - costo_operativo_hora
    resultados_ebitda.append(ebitda_simulado)

df_resultados = pd.DataFrame(resultados_ebitda, columns=['EBITDA_Proyectado'])
print(df_resultados.describe())</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por acuerdos de confidencialidad, este código es una representación simplificada del modelo lógico y los valores han sido modificados.
                    </div>
                </div>

                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Simulador Interactivo VDT (D3.js)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Haz clic en los nodos para expandir o colapsar el árbol. Esta visualización demuestra cómo los tiempos operativos base se agrupan matemáticamente para formar el costo de la flota y, finalmente, el EBITDA.</p>
                    
                    <div style="position: relative; width: 100%; aspect-ratio: 16/11; background: #0d1117; border-radius: 12px; border: 1px solid #30363d; overflow: hidden;">
                        <div id="p3-vdt-canvas" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;"></div>
                    </div>
                    
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por políticas de privacidad, este simulador utiliza cálculos aleatorios y no expone el estado financiero real de la empresa.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                <h4>Resultados del Proyecto</h4>
                Logramos transformar un mar de datos aislados en una herramienta gerencial directa:
                <ul>
                    <li><strong>Enfoque Directo al Problema:</strong> El gráfico de tornado nos mostró con total claridad cuál era la variable que más encarecía el costo de la mina. Dejamos de discutir sobre múltiples factores y nos enfocamos en el problema real.</li>
                    <li><strong>Control de Fugas de Capital:</strong> Al identificar cómo un pequeño aumento de segundos en los tiempos operativos se traducía en millones de dólares en gasto, la gerencia pudo crear planes de acción inmediatos.</li>
                    <li><strong>Escalabilidad del Modelo:</strong> Diseñé el árbol para que pudiera crecer; a medida que la operación lo requería, podíamos añadir nuevas variables al modelo y ver cómo afectaban el resultado global sin tener que reescribir todo desde cero.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">3. Cálculo de Correlación para el Tornado</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Para saber qué variable afectaba más, escribí esta lógica que evalúa matemáticamente la correlación entre cada rama pequeña del árbol (los tiempos en segundos) y el resultado financiero final.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            sensibilidad_tornado.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.8rem; line-height: 1.4;"><code>import pandas as pd
import scipy.stats as stats
import matplotlib.pyplot as plt

def generar_datos_tornado(df_simulacion, variable_objetivo='EBITDA'):
    correlaciones = {}
    columnas_operativas = [col for col in df_simulacion.columns if col != variable_objetivo]
    
    for col in columnas_operativas:
        coef, _ = stats.spearmanr(df_simulacion[col], df_simulacion[variable_objetivo])
        correlaciones[col] = coef
        
    df_tornado = pd.DataFrame.from_dict(correlaciones, orient='index', columns=['Impacto'])
    df_tornado['Impacto_Absoluto'] = df_tornado['Impacto'].abs()
    df_tornado = df_tornado.sort_values(by='Impacto_Absoluto', ascending=True)
    
    return df_tornado

datos_impacto = generar_datos_tornado(df_simulacion_completa)
print("Variables ordenadas para Graficar en Tornado:")
print(datos_impacto)</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Fragmento representativo del análisis de correlación usado para ordenar el gráfico de tornado final.
                    </div>
                </div>

                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">4. Gráfico de Tornado (Impacto en EBITDA)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">El producto final del informe gerencial. Pasa el cursor por las barras para ver cómo las variables operativas encarecen o abaratan el costo general de la operación.</p>
                    
                    <div id="tornado-chart" style="width: 100%; background: radial-gradient(circle at center, #0b1a2b, #020617); border-radius: 12px; border: 1px solid #1e293b; overflow: hidden; display: flex; justify-content: center; position: relative;"></div>
                    
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Nota: Por políticas de privacidad, este gráfico genera datos aleatorios y no expone sensibilidades financieras reales de la compañía.
                    </div>
                </div>

            </div>`
    },
    en: {
        title: "Value Driver Tree & EBITDA Simulation",
        shortDesc: "Stochastic Model to Identify Capital Leakage",
        company: "International Project Management (Cerrejón)",
        sector: "Financial Planning / Management Control",
        tools: "Excel, @Risk, Python, Financial Modeling",
        outcome: "Identification of the Main Cost Driver",
        body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                In mine financial management, understanding which operational factor most affects our profits (EBITDA) is a huge challenge because everything is connected. In this project, I was tasked with building a management report based on a Value Driver Tree model. The idea of this tree is that it starts with the final financial result and branches out into smaller and smaller elements, until it reaches specific operational variables, like the exact seconds a machine takes to complete a task.<br><br>

                The challenge was that using fixed averages for planning did not reflect the reality of the mine. To solve this, I used a tool called @Risk. What I did was take each operational variable in the tree (like real times in seconds) and set it to vary between a maximum and a minimum. Then, the model automatically calculated how that variability altered the results of all the other branches upwards, all the way to the organization's total EBITDA.<br><br>

                This way, we stopped guessing which part of the operation was costing us the most money and transitioned to a mathematical model that told us exactly where we were losing profitability.<br><br>

                <h4>What I Did</h4>
                <ul>
                    <li><strong>Tree Modeling:</strong> I structured a financial model that connected on-the-ground operational variables (cycle times, delays in seconds) with the mine's profit and loss statements.</li>
                    <li><strong>Variability Simulation (@Risk):</strong> I assigned real ranges (minimums and maximums in seconds) to the base variables and ran thousands of iterations to simulate the daily behavior of the operation.</li>
                    <li><strong>Sensitivity Analysis:</strong> I extracted the results into an interactive tornado chart that clearly showed management exactly which variable was most affecting the company's expenses.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">1. Stochastic Simulation Model</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">This script replicates the logic I used to inject variability between maximums and minimums (in seconds) into the base branches of the tree, calculating the final impact on EBITDA through thousands of iterations.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            vdt_montecarlo_ebitda.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.75rem; line-height: 1.4;"><code>import numpy as np
import pandas as pd

operational_variables = {
    'load_time_sec': (120, 145, 180),
    'haul_time_sec': (600, 650, 800),
    'operational_delays_sec': (45, 60, 120)
}

iterations = 10000
ebitda_results = []

for i in range(iterations):
    t_load = np.random.triangular(*operational_variables['load_time_sec'])
    t_haul = np.random.triangular(*operational_variables['haul_time_sec'])
    delays = np.random.triangular(*operational_variables['operational_delays_sec'])
    
    total_cycle_time = t_load + t_haul + delays
    trips_per_hour = 3600 / total_cycle_time
    
    tons_per_trip = 320
    price_per_ton = 115
    hourly_operating_cost = 450
    
    revenue = trips_per_hour * tons_per_trip * price_per_ton
    simulated_ebitda = revenue - hourly_operating_cost
    ebitda_results.append(simulated_ebitda)

df_results = pd.DataFrame(ebitda_results, columns=['Projected_EBITDA'])
print(df_results.describe())</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to confidentiality agreements, this code is a representation of the logical model, and the values have been altered.
                    </div>
                </div>

                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">2. Interactive VDT Simulator (D3.js)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">Click on the nodes to expand or collapse the tree. This visualization demonstrates how base operational times (cycle times) are mathematically grouped to form fleet costs and, ultimately, EBITDA.</p>
                    
                    <div style="position: relative; width: 100%; aspect-ratio: 16/11; background: #0d1117; border-radius: 12px; border: 1px solid #30363d; overflow: hidden;">
                        <div id="p3-vdt-canvas" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;"></div>
                    </div>
                    
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to privacy policies, this simulator uses randomized calculations and does not expose the actual financial state of the company.
                    </div>
                </div>
            </div>

            <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                <h4>Project Results</h4>
                We managed to transform a sea of isolated data into a direct management tool:
                <ul>
                    <li><strong>Direct Focus on the Problem:</strong> The tornado chart showed us with total clarity which variable was increasing the mine's costs the most. We stopped arguing about multiple factors and focused on the real problem.</li>
                    <li><strong>Capital Leakage Control:</strong> By identifying how a small increase of seconds in operational times translated into millions of dollars in expenses, management was able to create immediate action plans.</li>
                    <li><strong>Model Scalability:</strong> I designed the tree so it could grow; as the operation required it, we could add new variables to the model and see how they affected the overall result without having to rewrite everything from scratch.</li>
                </ul>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 400px), 1fr)); gap: 30px; margin: 40px 0; align-items: start;">
                
                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">3. Correlation Calculation for the Tornado</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">To know which variable had the most impact, I wrote this logic that mathematically evaluates the correlation between each small branch of the tree (the times in seconds) and the final financial result.</p>
                    
                    <div style="background: #0d1117; border: 1px solid #30363d; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                        <div style="background: #161b22; padding: 10px 15px; border-bottom: 1px solid #30363d; color: #8b949e; font-family: monospace; font-size: 0.85rem; display: flex; align-items: center;">
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ff5f56; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#ffbd2e; margin-right:8px;"></span>
                            <span style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#27c93f; margin-right:15px;"></span>
                            tornado_sensitivity.py
                        </div>
                        <pre style="margin: 0; padding: 20px; overflow-x: auto; color: #c9d1d9; font-family: 'Consolas', 'Courier New', monospace; font-size: 0.8rem; line-height: 1.4;"><code>import pandas as pd
import scipy.stats as stats
import matplotlib.pyplot as plt

def generar_datos_tornado(df_simulacion, variable_objetivo='EBITDA'):
    correlaciones = {}
    columnas_operativas = [col for col in df_simulacion.columns if col != variable_objetivo]
    
    for col in columnas_operativas:
        coef, _ = stats.spearmanr(df_simulacion[col], df_simulacion[variable_objetivo])
        correlaciones[col] = coef
        
    df_tornado = pd.DataFrame.from_dict(correlaciones, orient='index', columns=['Impact'])
    df_tornado['Absolute_Impact'] = df_tornado['Impact'].abs()
    df_tornado = df_tornado.sort_values(by='Absolute_Impact', ascending=True)
    
    return df_tornado

datos_impacto = generar_datos_tornado(df_simulacion_completa)
print("Sorted variables for Tornado Plotting:")
print(datos_impacto)</code></pre>
                    </div>
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Representative snippet of the correlation analysis used to order the final tornado chart.
                    </div>
                </div>

                <div style="min-width: 0;">
                    <h5 style="color: var(--text-main); font-size: 1rem; margin: 0 0 8px 0;">4. Sensitivity Analysis (Tornado)</h5>
                    <p style="color: var(--text-main); opacity: 0.85; font-size: 0.85rem; margin: 0 0 15px 0; line-height: 1.6; text-align: justify;">The final product of the management report. Hover over the bars to see how operational variables make the overall cost of the operation more expensive or cheaper.</p>
                    
                    <div id="tornado-chart" style="width: 100%; background: radial-gradient(circle at center, #0b1a2b, #020617); border-radius: 12px; border: 1px solid #1e293b; overflow: hidden; display: flex; justify-content: center; position: relative;"></div>
                    
                    <div style="margin-top: 10px; font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center;">
                        * Note: Due to privacy policies, this chart generates random data and does not expose real business sensitivities.
                    </div>
                </div>

            </div>`
    },
    gradient: "url('assets/images/proyecto3.png')"
}
    /*
        // PROYECTO 4: Personal
        {
            id: "p4",
            es: {
                title: "App RioPet",
                shortDesc: "Localización de Animales",
                company: "Proyecto Personal / Startup",
                sector: "Tecnología / Impacto Social",
                tools: "Por definir (Arquitectura)",
                outcome: "Fase Conceptual",
                body: "Diseño conceptual y arquitectura de una plataforma de impacto social enfocada en el rastreo y localización de animales perdidos en Riohacha. El proyecto involucra la estructuración de la base de datos geoespacial y la planificación de la lógica backend para un emparejamiento eficiente de reportes."
            },
            en: {
                title: "RioPet App",
                shortDesc: "Animal Localization Tracking",
                company: "Personal Project / Startup",
                sector: "Technology / Social Impact",
                tools: "TBD (Architecture Phase)",
                outcome: "Conceptual Phase",
                body: "Conceptual design and architecture of a social impact platform focused on tracking and locating lost animals in Riohacha. The project involves structuring the geospatial database and planning the backend logic for efficient report matching."
            },
            gradient: "url('assets/images/proyecto4.png')"
        },
    
        // PROYECTO 5: Personal
        {
            id: "p5",
            es: {
                title: "App Cursario",
                shortDesc: "Cursos Para Exámenes de Concursos",
                company: "Proyecto Personal / Emprendimiento",
                sector: "EdTech (Educación)",
                tools: "React, Js, Firebase",
                outcome: "Fase Conceptual",
                body: "Ideación de una plataforma educativa (EdTech) diseñada para estructurar rutas de aprendizaje y simulacros dirigidos a aspirantes de concursos públicos. El enfoque principal radica en el modelado de datos para gestionar perfiles de usuarios, progreso académico y bancos de preguntas dinámicos."
            },
            en: {
                title: "Cursario App",
                shortDesc: "Courses for Civil Service Exams",
                company: "Personal Project / Startup",
                sector: "EdTech (Education)",
                tools: "React, Js, Firebase",
                outcome: "Conceptual Phase",
                body: "Ideation of an educational platform (EdTech) designed to structure learning paths and mock tests aimed at civil service exam candidates. The main focus lies in data modeling to manage user profiles, academic progress, and dynamic question banks."
            },
            gradient: "url('assets/images/proyecto5.png')"
        }
            */
];

const writingData = [
    { 
        id: "w1", 
        es: { 
            title: "Transfiere archivos de iOS a Windows vía Wi-Fi", 
            shortDesc: "Configura una carpeta compartida (SMB) para pasar Gigas de fotos y videos a máxima velocidad.", 
            client: "Infraestructura / Productividad", 
            tools: "Windows, iOS, Red Local", 
            outcome: "Lectura: 4 min", 
            body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                <p>Cualquier persona que haya intentado hacer una copia de seguridad de un iPhone conectándolo por cable a un PC con Windows conoce la frustración: desconexiones constantes, transferencias lentas y una gestión de carpetas (DCIM) que es un verdadero caos. El ecosistema cerrado de Apple no se lleva bien de forma nativa con el explorador de archivos de Microsoft.</p>
                
                <p>Sin embargo, no necesitamos cables ni pagar suscripciones en la nube para resolver esto. La solución más elegante, rápida y privada es utilizar los protocolos de red local que ya vienen integrados en ambos sistemas operativos: <strong>El protocolo SMB (Server Message Block)</strong>.</p>
                
                <p>En este artículo te muestro cómo crear tu propia nube local para transferir archivos entre iOS y Windows a la velocidad máxima de tu Wi-Fi.</p>

                <p>Antes de empezar, necesitamos que tanto el PC como el iPhone estén conectados a la misma red Wi-Fi.</p>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/WIFI1.png" alt="Red Wi-Fi en Windows" class="clickable-img" style="max-height: 350px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/WIFI2.jpg" alt="Red Wi-Fi en iOS" class="clickable-img" style="max-height: 200px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Nota: Haz clic en las imágenes para ampliarlas.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Paso 1: Preparar la Carpeta Compartida en Windows</h4>
                <p>El primer paso es decirle a Windows qué carpeta queremos hacer accesible desde la red.</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Crea una carpeta en tu PC (por ejemplo, en el Escritorio) y llámala "Backup".</li>
                    <li style="margin-bottom: 8px;">Haz clic derecho sobre ella y selecciona <strong>Propiedades</strong> > pestaña <strong>Compartir</strong> > <strong>Compartir...</strong>.</li>
                    <li style="margin-bottom: 8px;">En el menú desplegable selecciona <strong>Todos</strong> y haz clic en <strong>Agregar</strong>. Después, asegúrate de seleccionar el nivel de permiso <strong>Lectura y escritura</strong> para que puedas copiar y borrar archivos desde el celular.</li>
                    <li style="margin-bottom: 8px;">Por último, haz clic en <strong>Compartir</strong> y luego en <strong>Listo</strong>.</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/carpeta1.png" alt="Configuración de permisos en Windows" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/carpeta2.png" alt="Confirmación de carpeta compartida" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Nota: Haz clic en las imágenes para ampliarlas.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Paso 2: Descubrir tu Dirección IP Local</h4>
                <p>Tu celular necesita saber exactamente a qué computadora conectarse dentro de tu red Wi-Fi. Para esto, necesitamos la dirección IP de tu PC.</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Abre el menú inicio de Windows, escribe <strong>cmd</strong> y presiona Enter.</li>
                    <li style="margin-bottom: 8px;">En la consola, escribe <code>ipconfig</code> y presiona Enter.</li>
                    <li style="margin-bottom: 8px;">Busca la línea que dice <strong>Dirección IPv4</strong>. Será un número parecido a <code>192.168.1.XXX</code>. Anótalo.</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/CMD1.png" alt="Abriendo el CMD en Windows" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/CMD2.png" alt="Obteniendo la dirección IPv4" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Nota: Haz clic en las imágenes para ampliarlas.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Paso 3: La Magia en tu iPhone / iPad</h4>
                <p>Ahora que la carpeta está compartida, vamos a conectarnos desde iOS usando la aplicación nativa "Archivos".</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Abre la app <strong>Archivos</strong> en tu iPhone.</li>
                    <li style="margin-bottom: 8px;">Toca el ícono de los tres puntos (...) en la esquina superior derecha y selecciona <strong>Conectarse al servidor</strong>.</li>
                    <li style="margin-bottom: 8px;">Escribe la IP que anotaste en el paso anterior (Ej: <code>192.168.1.XXX</code>).</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 10px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/connect.jpg" alt="Menú conectarse al servidor en iOS" class="clickable-img" style="max-height: 300px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 10px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/server_ip.jpg" alt="Ingresando la dirección IP" class="clickable-img" style="max-height: 300px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Nota: Haz clic en las imágenes para ampliarlas.
                </div>

                <p>Selecciona <strong>Usuario registrado</strong> e ingresa el nombre de usuario y la contraseña con los que inicias sesión en tu computadora. Normalmente, este es tu correo de Microsoft (o tu nombre de usuario local) y la contraseña asociada a ese correo o la que usas para desbloquear el PC.</p>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/name_pass.jpg" alt="Ingresando credenciales de Windows" class="clickable-img" style="max-height: 380px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/server.jpg" alt="Conexión exitosa a la carpeta" class="clickable-img" style="max-height: 380px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Nota: Haz clic en las imágenes para ampliarlas.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Conclusión</h4>
                <p>Y con esto ya sería todo: tu carpeta de Windows aparecerá en la app Archivos de iOS y podrás seleccionar cientos de fotos o videos pesados para copiarlos directamente allí. La transferencia se realizará vía Wi-Fi, sin pasar por los servidores de Apple o Microsoft, manteniendo total privacidad y aprovechando el 100% del ancho de banda de tu router local.</p>
            </div>` 
        }, 
        en: { 
            title: "Transfer Files from iOS to Windows via Wi-Fi", 
            shortDesc: "Set up a shared folder (SMB) to transfer Gigabytes of photos and videos at maximum speed.", 
            client: "Infrastructure / Productivity", 
            tools: "Windows, iOS, Local Network", 
            outcome: "4 min read", 
            body: `
            <div style="text-align: justify; line-height: 1.7; color: var(--text-main); margin-top: -45px;">
                <p>Anyone who has tried backing up an iPhone by plugging it into a Windows PC knows the frustration: constant disconnections, slow transfers, and a folder structure (DCIM) that is absolute chaos. Apple's closed ecosystem doesn't play nicely with Microsoft's file explorer out of the box.</p>
                
                <p>However, we don't need cables or cloud subscriptions to solve this. The most elegant, fast, and private solution is to use local network protocols that are already built into both operating systems: <strong>The SMB (Server Message Block) protocol</strong>.</p>
                
                <p>In this article, I'll show you how to create your own local cloud to transfer files between iOS and Windows at your Wi-Fi's maximum speed.</p>

                <p>Before we start, we need both the PC and the iPhone to be connected to the same Wi-Fi network.</p>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/WIFI1.png" alt="Wi-Fi network on Windows" class="clickable-img" style="max-height: 350px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/WIFI2.jpg" alt="Wi-Fi network on iOS" class="clickable-img" style="max-height: 200px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Note: Click on the images to enlarge them.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Step 1: Prepare the Shared Folder in Windows</h4>
                <p>The first step is to tell Windows which folder we want to make accessible over the network.</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Create a folder on your PC (for example, on your Desktop) and name it "Backup".</li>
                    <li style="margin-bottom: 8px;">Right-click on it, select <strong>Properties</strong> > go to the <strong>Sharing</strong> tab > and click on <strong>Share...</strong>.</li>
                    <li style="margin-bottom: 8px;">In the dropdown menu, select <strong>Everyone</strong> and click <strong>Add</strong>. Then, make sure to set the Permission Level to <strong>Read/Write</strong> so you can copy and delete files from your phone.</li>
                    <li style="margin-bottom: 8px;">Finally, click on <strong>Share</strong> and then <strong>Done</strong>.</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/carpeta1.png" alt="Windows sharing permissions" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/carpeta2.png" alt="Shared folder confirmation" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Note: Click on the images to enlarge them.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Step 2: Discover Your Local IP Address</h4>
                <p>Your phone needs to know exactly which computer to connect to within your Wi-Fi network. For this, we need your PC's IP address.</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Open the Windows start menu, type <strong>cmd</strong>, and press Enter.</li>
                    <li style="margin-bottom: 8px;">In the command prompt, type <code>ipconfig</code> and press Enter.</li>
                    <li style="margin-bottom: 8px;">Look for the line that says <strong>IPv4 Address</strong>. It will be a number like <code>192.168.1.XXX</code>. Write it down.</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/CMD1.png" alt="Opening CMD in Windows" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/CMD2.png" alt="Getting the IPv4 address" class="clickable-img" style="max-height: 280px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Note: Click on the images to enlarge them.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Step 3: The Magic on your iPhone / iPad</h4>
                <p>Now that the folder is shared, let's connect from iOS using the native "Files" app.</p>
                <ul style="padding-left: 25px; margin-bottom: 20px;">
                    <li style="margin-bottom: 8px;">Open the <strong>Files</strong> app on your iPhone.</li>
                    <li style="margin-bottom: 8px;">Tap the three-dot icon (...) in the top right corner and select <strong>Connect to Server</strong>.</li>
                    <li style="margin-bottom: 8px;">Type the IP address you wrote down in the previous step (e.g., <code>192.168.1.XXX</code>).</li>
                </ul>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 10px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/connect.jpg" alt="Connect to server menu in iOS" class="clickable-img" style="max-height: 300px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 10px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/server_ip.jpg" alt="Entering the IP address" class="clickable-img" style="max-height: 300px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Note: Click on the images to enlarge them.
                </div>

                <p>Select <strong>Registered User</strong> and enter the username and password you use to log into your computer. Typically, this is your Microsoft email (or your local username) and the password associated with that account or the one you use to unlock your PC.</p>

                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin: 30px 0 10px 0;">
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/name_pass.jpg" alt="Entering Windows credentials" class="clickable-img" style="max-height: 380px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                    <div style="flex: 0.5 1 300px; background: #111; border: 1px solid #333; border-radius: 12px; padding: 20px; box-sizing: border-box; display: flex; justify-content: center; align-items: center;">
                        <img src="assets/images/server.jpg" alt="Successful connection to the folder" class="clickable-img" style="max-height: 380px; width: auto; max-width: 100%; border-radius: 8px; display: block;">
                    </div>
                </div>
                <div style="font-size: 0.8rem; color: var(--text-main); opacity: 0.7; font-style: italic; text-align: center; margin-bottom: 30px;">
                    * Note: Click on the images to enlarge them.
                </div>

                <h4 style="color: var(--text-main); font-size: 1.2rem; margin: 30px 0 15px 0;">Conclusion</h4>
                <p>And that's it! Your Windows folder will now appear in the iOS Files app, and you can select hundreds of photos or heavy videos to copy them directly there. The transfer will happen over Wi-Fi, without going through Apple or Microsoft servers, maintaining total privacy and utilizing 100% of your local router's bandwidth.</p>
            </div>` 
        }, 
        gradient: "url('assets/images/iospc.png')" 
    }
];

const servicesData = [
    { 
        id: "srv1", 
        es: { 
            title: "Análisis de Datos y Modelado", 
            shortDesc: "Extracción, limpieza y estadística aplicada.", 
            client: "Estrategia", 
            tools: "SQL, Python, R, Pandas", 
            outcome: "Decisiones fundamentadas", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>Transformo datos crudos, estructurados y desestructurados en activos estratégicos. Mi enfoque va más allá de generar un simple reporte; aplico el rigor de la ingeniería y la estadística para resolver problemas de negocio complejos y respaldar la toma de decisiones con alta precisión numérica.</p>
                    
                    <p>Utilizando lenguajes como <strong>SQL</strong>, <strong>Python</strong> y <strong>R</strong>, realizo la extracción, limpieza y transformación de grandes volúmenes de información. Me apoyo fuertemente en librerías como <strong>Pandas</strong> y <strong>NumPy</strong> para el preprocesamiento, la imputación de valores nulos y la estructuración de los datos, garantizando una base sólida y confiable antes de cualquier análisis.</p>
                    
                    <p>Posteriormente, ejecuto un Análisis Exploratorio de Datos (EDA) profundo para identificar anomalías y descubrir patrones ocultos. Finalmente, implemento modelos estadísticos que van desde regresiones hasta <strong>simulaciones probabilísticas (como Monte Carlo)</strong> para la evaluación de riesgos y la proyección precisa de escenarios futuros.</p>
                </div>
            ` 
        }, 
        en: { 
            title: "Data Analysis & Modeling", 
            shortDesc: "Extraction, cleaning, and applied statistics.", 
            client: "Strategy", 
            tools: "SQL, Python, R, Pandas", 
            outcome: "Informed decisions", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>I transform raw, structured, and unstructured data into strategic assets. My approach goes beyond generating a simple report; I apply the rigor of engineering and statistics to solve complex business problems and support decision-making with high numerical precision.</p>
                    
                    <p>Using languages such as <strong>SQL</strong>, <strong>Python</strong>, and <strong>R</strong>, I extract, clean, and transform large volumes of information. I rely heavily on libraries like <strong>Pandas</strong> and <strong>NumPy</strong> for preprocessing, null-value imputation, and data structuring, ensuring a solid and reliable foundation before any analysis.</p>
                    
                    <p>Subsequently, I execute deep Exploratory Data Analysis (EDA) to identify anomalies and uncover hidden patterns. Finally, I implement statistical models ranging from regressions to <strong>probabilistic simulations (like Monte Carlo)</strong> for risk assessment and the accurate forecasting of future scenarios.</p>
                </div>
            ` 
        }, 
        gradient: "url('assets/images/Analisis_Datos.png')"
    },
    { 
        id: "srv2", 
        es: { 
            title: "Programación y RPA", 
            shortDesc: "Automatización de flujos de trabajo.", 
            client: "Operaciones", 
            tools: "Python, Java, Web Scraping", 
            outcome: "Eficiencia y ahorro de tiempo", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>Escribo código estructurado y eficiente para resolver cuellos de botella operativos y transformar procesos manuales propensos a errores en flujos de trabajo corporativos altamente escalables.</p>
                    
                    <p>Desarrollo scripts personalizados utilizando <strong>Python</strong> y <strong>Java</strong> para interactuar con sistemas, manipular archivos locales y gestionar bases de datos. Además, diseño soluciones de <strong>Web Scraping</strong> (mediante librerías como BeautifulSoup o Selenium) para extraer y procesar automáticamente grandes volúmenes de datos desde portales web o APIs de terceros.</p>
                    
                    <p>Al implementar esta <strong>Automatización Robótica de Procesos (RPA)</strong>, logro reducir drásticamente los tiempos de ejecución y elimino las tareas repetitivas, permitiendo que los equipos de trabajo liberen su agenda y se enfoquen en el análisis real y la estrategia de negocio.</p>
                </div>
            ` 
        }, 
        en: { 
            title: "Programming & RPA", 
            shortDesc: "Workflow automation.", 
            client: "Operations", 
            tools: "Python, Java, Web Scraping", 
            outcome: "Efficiency and time-saving", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>I write structured and efficient code to solve operational bottlenecks and transform error-prone manual processes into highly scalable corporate workflows.</p>
                    
                    <p>I develop custom scripts using <strong>Python</strong> and <strong>Java</strong> to interact with systems, manipulate local files, and manage databases. Additionally, I design <strong>Web Scraping</strong> solutions (using libraries like BeautifulSoup or Selenium) to automatically extract and process large volumes of data from web portals or third-party APIs.</p>
                    
                    <p>By implementing this <strong>Robotic Process Automation (RPA)</strong>, I drastically reduce execution times and eliminate repetitive tasks, allowing teams to free up their schedules and focus on actual analysis and business strategy.</p>
                </div>
            ` 
        }, 
        gradient: "url('assets/images/Programacion.png')" 
    },
    { 
        id: "srv3", 
        es: { 
            title: "Creación de KPI's", 
            shortDesc: "Diseño de métricas de rendimiento.", 
            client: "Analítica", 
            tools: "Estrategia de Datos, SQL", 
            outcome: "Medición exacta", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>Un dato por sí solo no tiene valor si no está alineado con un objetivo. Mi enfoque consiste en diseñar e implementar Indicadores Clave de Rendimiento (KPIs) que midan de forma exacta la salud y el progreso de la empresa frente a sus metas estratégicas.</p>
                    
                    <p>Trabajo directamente con las bases de datos mediante <strong>consultas SQL optimizadas</strong> para definir la lógica matemática detrás de cada métrica (numeradores, denominadores, filtros y periodicidad). Me aseguro de que cada indicador sea medible, alcanzable y, sobre todo, accionable.</p>
                    
                    <p>No me limito a estructurar el número; aporto la <strong>interpretación analítica</strong> necesaria para definir umbrales de éxito y traducir resultados complejos en un lenguaje corporativo claro, facilitando intervenciones rápidas por parte de la gerencia.</p>
                </div>
            ` 
        }, 
        en: { 
            title: "KPI Creation", 
            shortDesc: "Performance metrics design.", 
            client: "Analytics", 
            tools: "Data Strategy, SQL", 
            outcome: "Accurate measurement", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>Data alone has no value if it is not aligned with an objective. My approach involves designing and implementing Key Performance Indicators (KPIs) that accurately measure the health and progress of the company against its strategic goals.</p>
                    
                    <p>I work directly with databases using <strong>optimized SQL queries</strong> to define the mathematical logic behind each metric (numerators, denominators, filters, and periodicity). I ensure that every indicator is measurable, achievable, and above all, actionable.</p>
                    
                    <p>I don't just structure the number; I provide the necessary <strong>analytical interpretation</strong> to set success thresholds and translate complex results into clear corporate language, facilitating rapid interventions by management.</p>
                </div>
            ` 
        }, 
        gradient: "url('assets/images/kpis.jpg')" 
    },
    { 
        id: "srv4", 
        es: { 
            title: "Dashboards Dinámicos", 
            shortDesc: "Visualización interactiva de alto impacto.", 
            client: "Visualización", 
            tools: "Power BI, Excel, DAX", 
            outcome: "Monitoreo en vivo", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>Construyo tableros de control interactivos que no solo muestran gráficos, sino que cuentan la historia detrás de los datos. Mi objetivo es crear interfaces visualmente intuitivas para que perfiles técnicos y no técnicos puedan explorar la información con total libertad.</p>
                    
                    <p>Utilizo herramientas líderes en la industria como <strong>Power BI</strong> y <strong>Excel Avanzado</strong>. El proceso comienza con un sólido modelado de datos (ETL), conectando múltiples fuentes mediante <strong>Power Query</strong> y estructurando relaciones eficientes. Luego, implemento cálculos complejos e inteligencia de tiempo utilizando lenguaje <strong>DAX</strong>.</p>
                    
                    <p>El resultado final es un entorno dinámico con segmentadores, tooltips y jerarquías interactivas (drill-downs), diseñado con principios de UI/UX para que la gerencia pueda monitorear la operación en tiempo real y tomar decisiones ágiles con un solo vistazo.</p>
                </div>
            ` 
        }, 
        en: { 
            title: "Dynamic Dashboards", 
            shortDesc: "High-impact interactive visualization.", 
            client: "Visualization", 
            tools: "Power BI, Excel, DAX", 
            outcome: "Live monitoring", 
            body: `
                <div style="text-align: justify; line-height: 1.7; color: var(--text-main);">
                    <p>I build interactive control panels that don't just show charts, but tell the story behind the data. My goal is to create visually intuitive interfaces so that both technical and non-technical profiles can explore the information with total freedom.</p>
                    
                    <p>I use industry-leading tools like <strong>Power BI</strong> and <strong>Advanced Excel</strong>. The process begins with solid data modeling (ETL), connecting multiple sources via <strong>Power Query</strong> and structuring efficient relationships. Then, I implement complex calculations and time intelligence using the <strong>DAX</strong> language.</p>
                    
                    <p>The final result is a dynamic environment with slicers, tooltips, and interactive hierarchies (drill-downs), designed with UI/UX principles so that management can monitor the operation in real-time and make agile decisions at a single glance.</p>
                </div>
            ` 
        }, 
        gradient: "url('assets/images/Dashboard.png')" 
    }
];

const translations = {
    es: {
        role: "Analista de Datos e Ingeniero de Sistemas",
        nav_home: "Inicio", nav_about: "Sobre mí", nav_projects: "Proyectos", nav_writing: "Artículos", nav_stack: "Herramientas", nav_services: "Servicios", nav_contact: "Contacto",
        lang_btn_desktop: "US", lang_btn_mobile: "US", theme_light_desktop: "☀️ Claro", theme_light_mobile: "☀️", theme_dark_desktop: "🌙 Oscuro", theme_dark_mobile: "🌙",
        home_title: "Transformando datos en <br>decisiones estratégicas.", home_desc: "Ingeniero de Sistemas especializado en automatización de procesos operativos, análisis exploratorio con Python y visualización avanzada.",
        about_title: "Sobre Mí",
        proj_subtitle: "Proyectos", writing_subtitle: "Publicaciones Recientes", meta_client: "Contexto", meta_company: "Empresa", meta_sector: "Sector", meta_tools: "Herramientas", meta_outcome: "Info",
        stack_title: "Mi Stack Tecnológico", stack_desc: "Herramientas y lenguajes que utilizo en mi día a día para construir soluciones.", stack_cat1: "Análisis & Backend", stack_cat2: "Visualización & Datos", stack_cat3: "Infraestructura & Nube",
        services_title: "Servicios", services_desc: "Áreas de especialidad técnica enfocadas en resolver problemas corporativos.", srv1_title: "Análisis de Datos", srv1_desc: "Extracción, limpieza y transformación de datos complejos usando SQL y Python.", srv2_title: "Automatización (RPA)", srv2_desc: "Desarrollo de scripts en VBA y Python para reducir horas de trabajo manual en tareas repetitivas.",
        page_home: "Inicio", page_about: "Sobre mí", page_projects: "Proyectos", page_writing: "Artículos", page_stack: "Herramientas", page_services: "Servicios", page_contact: "Contacto",

        status_available: "Disponible para nuevas oportunidades",
        home_contact_me: "Contáctame", home_view_cv: "Descargar CV", home_desc_bento: "Ingeniero de Sistemas de 26 años especializado en el análisis profundo de datos y el modelado estadístico. Me dedico a identificar patrones estratégicos y simular escenarios complejos, con el firme propósito de transformar grandes volúmenes de información en decisiones de negocio altamente precisas y fundamentadas.",
        read_more_about: "Leer más sobre mí →", explore_work: "Ver proyecto →", read_articles: "Leer artículo →",
        featured_p1_desc: "Simulación de costos con Python y Excel para optimizar la toma de decisiones.", featured_a1_desc: "Configura una carpeta compartida (SMB) para pasar fotos y videos a tu PC",

        about_page_title: "Sobre mí", about_page_subtitle: "Algunas cosas que deberías saber", about_label_intro: "Intro",
        about_intro_p1: "Hola, soy Aldahir Mendoza, Ingeniero de Sistemas. Me especializo en crear soluciones basadas en datos y desarrollar flujos de automatización que resuelven problemas complejos y optimizan la eficiencia operativa a nivel corporativo, utilizando herramientas de software especializadas en el tratamiento de la información.",
        about_intro_p2: "Cuento con una sólida experiencia en análisis estadístico exploratorio, Web Scraping y el diseño de herramientas funcionales. Creo firmemente que un código backend limpio y escalable debe integrarse a la perfección con visualizaciones dinámicas, garantizando el rendimiento y facilitando la toma de decisiones estratégicas en cualquier área de una organización.",
        about_intro_p3: "Por fuera del mundo de las bases de datos y la programación, soy una persona muy orientada a la disciplina física. Me apasiona jugar al fútbol, salir a correr y el entrenamiento continuo, ya que considero que mantener un estilo de vida activo es clave para tener una mente ágil y el equilibrio perfecto para despejarse frente a los retos técnicos.",
        about_label_exp: "Experiencia",
        exp_date_1: "Jul 2023 — May 2025", exp_desc_1: "Analista de Negocios ll en IPM",
        exp_date_2: "Oct 2022 — Dic 2022", exp_desc_2: "Trainee CPITS en Trend Micro",
        exp_date_3: "Feb 2022 — Ago 2022", exp_desc_3: "Practicante en Cerrejón",

        about_label_edu: "Educación",
        edu_date_1: "Ene 2017 — Nov 2022", edu_desc_1: "Ingeniería de Sistemas e Informática en Universidad Pontificia Bolivariana",
        edu_date_2: "May 2021 — Dic 2021", edu_desc_2: "Fundamentos de Programación en Universidad Industrial de Santander",

        about_label_cert: "Certificados",
        cert_date_1: "2024", cert_desc_1: "Certificado en Python",
        cert_date_2: "2024", cert_desc_2: "Certificado de Estadística y Probabilidad",
        cert_date_3: "2021", cert_desc_3: "Certificado Fundamentos de Programación",

        stack_page_subtitle: "Software y herramientas que utilizo", stack_label_software: "Software", stack_cat_lang: "Lenguaje", stack_cat_ide: "IDE", stack_cat_db: "Base de Datos", stack_cat_data: "Datos / BI",

        contact_page_title: "Contacto", contact_page_subtitle: "Hablemos sobre trabajar juntos", contact_btn_phone: "Copiar teléfono", contact_btn_copy: "Copiar email", contact_btn_copied: "Copiado", contact_or: "o", contact_ph_name: "Tu nombre", contact_ph_email: "Tu correo", contact_ph_message: "Tu mensaje", contact_btn_send: "Enviar", contact_response_time: "Respondo en 1-2 horas"
    },
    en: {
        role: "Data Analyst & Systems Engineer",
        nav_home: "Home", nav_about: "About", nav_projects: "Work", nav_writing: "Writing", nav_stack: "Stack", nav_services: "Services", nav_contact: "Contact",
        lang_btn_desktop: "ES", lang_btn_mobile: "ES", theme_light_desktop: "☀️ Light", theme_light_mobile: "☀️", theme_dark_desktop: "🌙 Dark", theme_dark_mobile: "🌙",
        home_title: "Transforming data into <br>strategic decisions.", home_desc: "Systems Engineer specialized in operational process automation, exploratory analysis with Python, and advanced data visualization.",
        about_title: "About Me",
        proj_subtitle: "Projects", writing_subtitle: "Recent Publications", meta_client: "Context", meta_company: "Company", meta_sector: "Industry", meta_tools: "Tools Used", meta_outcome: "Info",
        stack_title: "Tech Stack", stack_desc: "Tools and languages I use daily to build solutions.", stack_cat1: "Analysis & Backend", stack_cat2: "Visualization & Data", stack_cat3: "Cloud & Infrastructure",
        services_title: "Services", services_desc: "Technical specialty areas focused on solving corporate challenges.", srv1_title: "Data Analysis", srv1_desc: "Extraction, cleaning, and transformation of complex data using SQL and Python.", srv2_title: "Automation (RPA)", srv2_desc: "Development of VBA and Python scripts to reduce manual work hours in repetitive tasks.",
        page_home: "Home", page_about: "About Me", page_projects: "Work", page_writing: "Writing", page_stack: "Tech Stack", page_services: "Services", page_contact: "Contact",

        status_available: "Available for new opportunities",
        home_contact_me: "Contact me", home_view_cv: "Download CV", home_desc_bento: "26-year-old Systems Engineer specializing in deep data analysis and statistical modeling. I focus on identifying strategic patterns and simulating complex scenarios, with the firm purpose of transforming large volumes of information into highly accurate and informed business decisions.",
        read_more_about: "Read more about me →", explore_work: "Explore work →", read_articles: "Read article →",
        featured_p1_desc: "Cost simulation with Python and Excel to optimize decision-making.", featured_a1_desc: "Set up a shared folder (SMB) to transfer photos and videos to your PC.",

        about_page_title: "About", about_page_subtitle: "A few things you should know", about_label_intro: "Intro",
        about_intro_p1: "Hello, I'm Aldahir Mendoza, a Systems Engineer. I specialize in creating data-driven solutions and developing automation workflows that solve complex problems and optimize operational efficiency at a corporate level, using specialized software tools for data processing.",
        about_intro_p2: "I have solid experience in exploratory statistical analysis, Web Scraping, and the design of functional tools. I firmly believe that clean, scalable backend code must integrate seamlessly with dynamic visualizations, ensuring performance and facilitating strategic decision-making in any area of an organization.",
        about_intro_p3: "Outside the world of databases and programming, I am highly oriented toward physical discipline. I am passionate about playing football, running, and continuous training, as I consider maintaining an active lifestyle key to keeping an agile mind and finding the perfect balance to clear my head when facing technical challenges.",
        about_label_exp: "Experience",
        exp_date_1: "Jul 2023 — May 2025", exp_desc_1: "Business Analyst ll at IPM",
        exp_date_2: "Oct 2022 — Dec 2022", exp_desc_2: "Trainee CPITS at Trend Micro",
        exp_date_3: "Feb 2022 — Aug 2022", exp_desc_3: "Internship at Cerrejón",

        about_label_edu: "Education",
        edu_date_1: "Jan 2017 — Nov 2022", edu_desc_1: "Bachelor of Systems and Informatics Engineering at Universidad Pontificia Bolivariana",
        edu_date_2: "May 2021 — Dec 2021", edu_desc_2: "Programming Fundamentals at Universidad Industrial de Santander",

        about_label_cert: "Certificates",
        cert_date_1: "2024", cert_desc_1: "Python Certificate",
        cert_date_2: "2024", cert_desc_2: "Statistics and Probability Certificate",
        cert_date_3: "2021", cert_desc_3: "Programming Fundamentals Certificate",

        stack_page_subtitle: "Software & physical products I use", stack_label_software: "Software", stack_cat_lang: "Language", stack_cat_ide: "IDE", stack_cat_db: "Database", stack_cat_data: "Data / BI",

        contact_page_title: "Contact", contact_page_subtitle: "Let's talk about working together", contact_btn_phone: "Copy phone", contact_btn_copy: "Copy email", contact_btn_copied: "Copied", contact_or: "or", contact_ph_name: "Your name", contact_ph_email: "Your email", contact_ph_message: "Your message", contact_btn_send: "Send", contact_response_time: "Around 1-2 hours to respond"
    }
};

let currentLang = localStorage.getItem('language') || 'es';
const savedTheme = localStorage.getItem('theme');

if (savedTheme === 'light') document.body.classList.add('light-mode');

let currentProjId = projectsData[0].id;
let currentWritingId = writingData[0].id;
let currentServiceId = servicesData[0].id;

let pageContext = 'home';
const bodyId = document.body.id;
if (bodyId === 'body-about') pageContext = 'about';
if (bodyId === 'body-projects') pageContext = 'projects';
if (bodyId === 'body-writing') pageContext = 'writing';
if (bodyId === 'body-stack') pageContext = 'stack';
if (bodyId === 'body-services') pageContext = 'services';
if (bodyId === 'body-contact') pageContext = 'contact';

const burgerBtn = document.getElementById('burgerBtn');
const mobileModal = document.getElementById('mobileModal');

function checkBurgerVisibility() {
    if ((pageContext === 'projects' || pageContext === 'writing' || pageContext === 'services') && window.innerWidth <= 850) {
        if (burgerBtn) burgerBtn.setAttribute('style', 'display: flex !important;');
    } else {
        if (burgerBtn) burgerBtn.setAttribute('style', 'display: none !important;');
    }
}

if (burgerBtn) {
    burgerBtn.addEventListener('click', () => {
        if (mobileModal.classList.contains('open')) closeMobileModal();
        else openMobileModal();
    });
}

function openMobileModal() {
    const container = document.getElementById('modalListContainer');
    const modalTitle = document.getElementById('modalTitle');
    if (!container || !modalTitle) return;

    container.innerHTML = '';

    if (pageContext === 'projects') {
        modalTitle.textContent = translations[currentLang].proj_subtitle;
        populateModalList(projectsData, currentProjId, (id) => { currentProjId = id; updateAllData(); closeMobileModal(); });
    } else if (pageContext === 'writing') {
        modalTitle.textContent = translations[currentLang].writing_subtitle;
        populateModalList(writingData, currentWritingId, (id) => { currentWritingId = id; updateAllData(); closeMobileModal(); });
    } else if (pageContext === 'services') {
        modalTitle.textContent = translations[currentLang].nav_services;
        populateModalList(servicesData, currentServiceId, (id) => { currentServiceId = id; updateAllData(); closeMobileModal(); });
    }
    mobileModal.classList.add('open');
}

function closeMobileModal() { if (mobileModal) mobileModal.classList.remove('open'); }

function populateModalList(dataArray, currentId, clickCallback) {
    const container = document.getElementById('modalListContainer');
    dataArray.forEach(item => {
        const data = item[currentLang];
        const div = document.createElement('div');
        div.className = `modal-list-item ${item.id === currentId ? 'active' : ''}`;
        div.innerHTML = `<span>${data.title}</span>`;
        div.addEventListener('click', () => clickCallback(item.id));
        container.appendChild(div);
    });
}

function renderDesktopList(dataArray, currentId, containerId, clickCallback) {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (window.innerWidth > 850) {
        container.innerHTML = '';
        dataArray.forEach(item => {
            const data = item[currentLang];
            const div = document.createElement('div');
            div.className = `list-item ${item.id === currentId ? 'active' : ''}`;
            div.innerHTML = `<h4>${data.title}</h4><p>${data.shortDesc}</p>`;
            div.addEventListener('click', () => clickCallback(item.id));
            container.appendChild(div);
        });
    }
}

function renderDetail(dataArray, currentId, areaId) {
    const area = document.getElementById(areaId);
    if (!area) return;
    const item = dataArray.find(i => i.id === currentId);
    const data = item[currentLang];
    const t = translations[currentLang];

    // Animaciones fluidas
    area.classList.remove('animate-fade-up', 'delay-2');
    area.classList.remove('fade-content');
    void area.offsetWidth;
    area.classList.add('fade-content');

    // Lógica para mostrar 4 datos en Proyectos y 3 en el resto
    let metaHtml = '';
    if (pageContext === 'projects') {
        metaHtml = `
            <div class="meta-item"><span>${t.meta_company}</span><strong>${data.company}</strong></div>
            <div class="meta-item"><span>${t.meta_sector}</span><strong>${data.sector}</strong></div>
            <div class="meta-item"><span>${t.meta_tools}</span><strong>${data.tools}</strong></div>
            <div class="meta-item"><span>${t.meta_outcome}</span><strong>${data.outcome}</strong></div>
        `;
    } else {
        metaHtml = `
            <div class="meta-item"><span>${t.meta_client}</span><strong>${data.client}</strong></div>
            <div class="meta-item"><span>${t.meta_tools}</span><strong>${data.tools}</strong></div>
            <div class="meta-item"><span>${t.meta_outcome}</span><strong>${data.outcome}</strong></div>
        `;
    }

    area.innerHTML = `
        <div class="hero-img" style="background: ${item.gradient}"></div>
        <h2 style="font-size: 2.2rem; margin-bottom: 10px;">${data.title}</h2>
        <div class="metadata-grid">
            ${metaHtml}
        </div>
        <p style="font-size: 1.05rem; color: var(--text-muted); line-height: 1.8; margin-bottom: 40px;">${data.body}</p>
        
        <footer class="site-footer" style="margin-top: auto; padding-top: 40px; border-top: 1px solid var(--border-color); text-align: center;">
            <p style="font-size: 0.8rem; margin: 0;">© 2026 RiverAxe by <a href="https://github.com/Aldamen" target="_blank">@aldamen</a></p>
        </footer>
    `;

    // ===============================================
    // 🔥 SOLUCIÓN AL SCROLL: VOLVER AL INICIO 🔥
    // ===============================================
    // Como tu layout usa flexbox, el scroll ocurre dentro de contenedores específicos.
    // Forzamos a ambos posibles contenedores a volver a la posición 0 (arriba).
    const splitDetailArea = document.querySelector('.split-detail-area');
    if (splitDetailArea) splitDetailArea.scrollTop = 0;

    const mainContentArea = document.querySelector('.main-content');
    if (mainContentArea) mainContentArea.scrollTop = 0;
    // ===============================================

    // Al final, intentamos iniciar el simulador por si acaso estamos en un proyecto con gráficos
    setTimeout(initSimulator, 100);
}

function updateAllData() {
    renderDesktopList(projectsData, currentProjId, 'projectListContainer', (id) => { currentProjId = id; updateAllData(); });
    renderDesktopList(writingData, currentWritingId, 'writingListContainer', (id) => { currentWritingId = id; updateAllData(); });
    renderDesktopList(servicesData, currentServiceId, 'serviceListContainer', (id) => { currentServiceId = id; updateAllData(); });
    renderDetail(projectsData, currentProjId, 'projectDetailArea');
    renderDetail(writingData, currentWritingId, 'writingDetailArea');
    renderDetail(servicesData, currentServiceId, 'serviceDetailArea');
}

const langBtn = document.getElementById('langBtn');
const themeBtn = document.getElementById('themeBtn');

function updateTexts() {
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (translations[currentLang][key]) {
            if (el.tagName === 'SPAN' && (el.parentElement.classList.contains('bottom-nav-item') || el.classList.contains('tooltip-text'))) {
                el.textContent = translations[currentLang][key];
            } else {
                el.innerHTML = translations[currentLang][key];
            }
        }
    });

    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        if (translations[currentLang][key]) el.setAttribute('placeholder', translations[currentLang][key]);
    });

    const mobileHeaderEl = document.querySelector('.mobile-header');
    if (mobileHeaderEl) mobileHeaderEl.textContent = translations[currentLang][`page_${pageContext}`];

    // ESTA ES LA VARIABLE QUE DETECTA EL MODO CLARO
    const isLight = document.body.classList.contains('light-mode');

    // ===============================================
    // 🔥 NUEVO: CAMBIAR EL LOGO SEGÚN EL TEMA 🔥
    // ===============================================
    const logos = document.querySelectorAll('.site-logo');
    logos.forEach(logo => {
        // Si isLight es verdadero pone el logo oscuro, si es falso pone el blanco
        logo.src = isLight ? 'assets/images/logonegro.png' : 'assets/images/logoblanco.png';
    });
    // ===============================================

    if (document.getElementById('langTextDesktop')) document.getElementById('langTextDesktop').innerHTML = translations[currentLang].lang_btn_desktop;
    if (document.getElementById('langTextMobile')) document.getElementById('langTextMobile').innerHTML = translations[currentLang].lang_btn_mobile;
    if (document.getElementById('themeTextDesktop')) document.getElementById('themeTextDesktop').innerHTML = isLight ? translations[currentLang].theme_dark_desktop : translations[currentLang].theme_light_desktop;
    if (document.getElementById('themeTextMobile')) document.getElementById('themeTextMobile').innerHTML = isLight ? translations[currentLang].theme_dark_mobile : translations[currentLang].theme_light_mobile;

    updateAllData();
}

if (langBtn) {
    langBtn.addEventListener('click', () => {
        currentLang = currentLang === 'es' ? 'en' : 'es';
        localStorage.setItem('language', currentLang);
        updateTexts();
    });
}

if (themeBtn) {
    themeBtn.addEventListener('click', () => {
        document.body.classList.add('theme-transition');
        document.body.classList.toggle('light-mode');
        localStorage.setItem('theme', document.body.classList.contains('light-mode') ? 'light' : 'dark');
        updateTexts();
    });
}

function centerBottomNav() {
    const bottomNav = document.querySelector('.bottom-nav-bar');
    const activeItem = document.querySelector('.bottom-nav-item.active');
    if (bottomNav && activeItem && window.innerWidth <= 850) {
        bottomNav.style.scrollBehavior = 'auto';
        const navRect = bottomNav.getBoundingClientRect();
        const itemRect = activeItem.getBoundingClientRect();
        const centerPos = bottomNav.scrollLeft + (itemRect.left - navRect.left) - (navRect.width / 2) + (itemRect.width / 2);
        bottomNav.scrollLeft = centerPos;
        setTimeout(() => { bottomNav.style.scrollBehavior = 'smooth'; }, 50);
    }
}

window.addEventListener('resize', () => {
    checkBurgerVisibility();
    updateAllData();
    closeMobileModal();
    centerBottomNav();
});

if (document.body.classList.contains('preload')) document.body.classList.remove('preload');
checkBurgerVisibility();
updateTexts();
window.addEventListener('load', centerBottomNav);
centerBottomNav();

// --- LÓGICA DE LA PÁGINA DE CONTACTO ---

// Funcionalidad de "Copiar Teléfono"
const copyPhoneBtn = document.getElementById('copyPhoneBtn');
if (copyPhoneBtn) {
    copyPhoneBtn.addEventListener('click', () => {
        // Pon aquí tu número de teléfono real
        const phoneToCopy = "+57 3175360228";

        navigator.clipboard.writeText(phoneToCopy).then(() => {
            const textSpan = document.getElementById('copyPhoneText');
            const iconSvg = document.getElementById('copyPhoneIcon');

            // Cambiar texto a "Copiado" y poner el icono del check
            textSpan.textContent = translations[currentLang].contact_btn_copied;
            iconSvg.innerHTML = '<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>';

            // Volver a la normalidad después de 2 segundos
            setTimeout(() => {
                textSpan.textContent = translations[currentLang].contact_btn_phone;
                iconSvg.innerHTML = '<path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/>';
            }, 2000);
        });
    });
}
const copyEmailBtn = document.getElementById('copyEmailBtn');
if (copyEmailBtn) {
    copyEmailBtn.addEventListener('click', () => {
        const emailToCopy = "lainmendoza08@hotmail.com";
        navigator.clipboard.writeText(emailToCopy).then(() => {
            const textSpan = document.getElementById('copyEmailText');
            const iconSvg = document.getElementById('copyEmailIcon');
            textSpan.textContent = translations[currentLang].contact_btn_copied;
            iconSvg.innerHTML = '<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>';
            setTimeout(() => {
                textSpan.textContent = translations[currentLang].contact_btn_copy;
                iconSvg.innerHTML = '<path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>';
            }, 2000);
        });
    });
}

const contactForm = document.getElementById('contactForm');
if (contactForm) {
    contactForm.addEventListener('submit', (e) => {
        e.preventDefault();
        
        const nameInput = document.getElementById('nameInput');
        const emailInput = document.getElementById('emailInput');
        const messageInput = document.getElementById('messageInput');
        const submitBtn = document.querySelector('.contact-submit-btn');
        
        let isValid = true;
        [nameInput, emailInput, messageInput].forEach(input => input.classList.remove('error'));

        if (!nameInput.value.trim()) { nameInput.classList.add('error'); isValid = false; }
        if (!emailInput.value.trim() || !emailInput.value.includes('@')) { emailInput.classList.add('error'); isValid = false; }
        if (!messageInput.value.trim()) { messageInput.classList.add('error'); isValid = false; }

        if (isValid) {
            // 1. Guardamos el texto original del botón y mostramos que está cargando
            const originalBtnText = submitBtn.textContent;
            submitBtn.textContent = currentLang === 'es' ? 'Enviando...' : 'Sending...';
            submitBtn.style.opacity = '0.7';
            submitBtn.style.pointerEvents = 'none'; // Evita doble clic

            // 2. Enviamos los datos a Formspree usando Fetch API
            // ⚠️ REEMPLAZA ESTA URL POR LA QUE TE DIO FORMSPREE ⚠️
            fetch('https://formspree.io/f/mzbqvqke', {
                method: 'POST',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    name: nameInput.value,
                    email: emailInput.value,
                    message: messageInput.value
                })
            }).then(response => {
                if (response.ok) {
                    alert(currentLang === 'es' ? "¡Mensaje enviado con éxito! Te responderé pronto." : "Message sent successfully! I'll reply soon.");
                    contactForm.reset();
                } else {
                    alert(currentLang === 'es' ? "Hubo un error al enviar el mensaje." : "There was an error sending the message.");
                }
            }).catch(error => {
                alert(currentLang === 'es' ? "Error de red. Revisa tu conexión." : "Network error. Please check your connection.");
            }).finally(() => {
                // 3. Restauramos el botón a la normalidad
                submitBtn.textContent = originalBtnText;
                submitBtn.style.opacity = '1';
                submitBtn.style.pointerEvents = 'auto';
            });
        }
    });

    document.querySelectorAll('.contact-form input, .contact-form textarea').forEach(input => {
        input.addEventListener('input', () => input.classList.remove('error'));
    });
}

let chartInstance = null;

function initSimulator() {
    const ctx = document.getElementById('simChart').getContext('2d');

    // 🎯 Generar datos "reales" tipo histograma (simulación operativa)
    function generateRealHistogram() {
        const data = [];
        const mean = 65;
        const std = 12;

        for (let i = 0; i <= 100; i++) {
            const val = Math.exp(-0.5 * Math.pow((i - mean) / std, 2)) * 26;
            const noise = (Math.random() - 0.5) * 3;
            data.push(Math.max(0, val + noise));
        }
        return data;
    }

    const realData = generateRealHistogram();

    // 📈 Distribución normal (curva roja)
    function normalDist(x, mean, std) {
        return Math.exp(-0.5 * Math.pow((x - mean) / std, 2)) * 26;
    }

    // 🎯 Plugin para línea vertical de la media
    const meanLinePlugin = {
        id: 'meanLine',
        afterDraw(chart) {
            const { ctx, scales: { x, y } } = chart;
            const mean = chart.config.data.datasets[1].meanValue;

            if (!mean) return;

            const xPos = x.getPixelForValue(mean);

            ctx.save();
            ctx.beginPath();
            ctx.setLineDash([5, 5]);
            ctx.strokeStyle = 'rgba(255,255,255,0.5)';
            ctx.moveTo(xPos, y.top);
            ctx.lineTo(xPos, y.bottom);
            ctx.stroke();
            ctx.restore();
        }
    };

    if (chartInstance) chartInstance.destroy();

    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Array.from({ length: 101 }, (_, i) => i),
            datasets: [
                {
                    label: 'Histograma',
                    data: realData,
                    backgroundColor: 'rgba(120,160,255,0.5)',
                    borderRadius: 2
                },
                {
                    label: 'Curva',
                    type: 'line',
                    data: [],
                    borderColor: '#ff4d4d',
                    borderWidth: 3,
                    pointRadius: 0,
                    tension: 0.4,
                    meanValue: 65
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    min: 0,
                    max: 26,
                    ticks: {
                        stepSize: 2,
                        color: '#888'
                    },
                    title: {
                        display: true,
                        text: 'Frecuencia de Eventos',
                        color: '#aaa'
                    },
                    grid: { color: '#222' }
                },
                x: {
                    min: 0,
                    max: 100,
                    ticks: {
                        stepSize: 10,
                        color: '#888'
                    },
                    title: {
                        display: true,
                        text: 'Valor de la Variable (Rendimiento)',
                        color: '#aaa'
                    },
                    grid: { display: false }
                }
            }
        },
        plugins: [meanLinePlugin]
    });

    const iMean = document.getElementById('input-mean');
    const iStd = document.getElementById('input-std');
    const msg = document.getElementById('sim-msg');

    function update() {
        const m = parseFloat(iMean.value);
        const s = parseFloat(iStd.value);

        document.getElementById('txt-media').innerText = m.toFixed(1);
        document.getElementById('txt-std').innerText = s.toFixed(1);

        // 🔴 Curva
        chartInstance.data.datasets[1].data =
            chartInstance.data.labels.map(x => normalDist(x, m, s));

        chartInstance.data.datasets[1].meanValue = m;

        // 🎯 Precisión (comparación con real)
        const errorMean = Math.abs(65 - m);
        const errorStd = Math.abs(12 - s);

        const accuracy = Math.max(0, 100 - (errorMean * 1.2 + errorStd * 2)).toFixed(0);

        const accSpan = document.getElementById('txt-acc');
        accSpan.innerText = accuracy + "%";

        // 🧠 Comentarios inteligentes
        if (accuracy >= 90) {
            msg.innerHTML = "Buen ajuste. Los parámetros están cerca de los valores operativos reales.";
            msg.style.color = "#4ade80";
            msg.style.background = "rgba(74,222,128,0.1)";
            accSpan.style.color = "#4ade80";

        } else if (accuracy >= 70) {
            msg.innerHTML = "Ajuste aceptable. Se recomienda afinar la desviación estándar.";
            msg.style.color = "#facc15";
            msg.style.background = "rgba(250,204,21,0.1)";
            accSpan.style.color = "#facc15";

        } else if (accuracy >= 40) {
            msg.innerHTML = "Ajuste bajo. La media o la volatilidad no reflejan bien la operación.";
            msg.style.color = "#fb923c";
            msg.style.background = "rgba(251,146,60,0.1)";
            accSpan.style.color = "#fb923c";

        } else {
            msg.innerHTML = "Ajuste incorrecto. Se requiere recalibración completa del modelo.";
            msg.style.color = "#ff4d4d";
            msg.style.background = "rgba(255,77,77,0.1)";
            accSpan.style.color = "#ff4d4d";
        }

        chartInstance.update();
    }

    iMean.addEventListener('input', update);
    iStd.addEventListener('input', update);

    update();


    

// ==========================================
// LÓGICA DEL SIMULADOR VDT (PROYECTO 3) - D3.js MEJORADO
// ==========================================

function renderVDT() {
    const containerNode = document.getElementById("p3-vdt-canvas");
    if (!containerNode) return;
    
    if (containerNode.querySelector("svg")) return;

    // 🎲 Generador de tiempos aleatorios (ej: entre 1.00 y 5.00 minutos)
    const randTime = () => (Math.random() * (5.0 - 1.0) + 1.0);
    // 🎲 Generador de costo de gasolina (ej: entre $2,000 y $5,000)
    const randGas = () => Math.floor(Math.random() * (5000 - 2000) + 2000);

    // 🌳 Estructura de Datos con TIPOS (time o currency)
    const data = {
        name: "EBITDA",
        type: "currency",
        children: [
            {
                name: "Costo Camiones",
                type: "currency",
                children: [
                    { name: "Tiempo 1", value: randTime(), type: "time" },
                    { name: "Tiempo 2", value: randTime(), type: "time" },
                    { name: "Tiempo 3", value: randTime(), type: "time" },
                    { name: "Tiempo 4", value: randTime(), type: "time" }
                ]
            },
            {
                name: "Costo Palas",
                type: "currency",
                children: [
                    { name: "Tiempo 1", value: randTime(), type: "time" },
                    { name: "Tiempo 2", value: randTime(), type: "time" },
                    { name: "Tiempo 3", value: randTime(), type: "time" },
                    { name: "Tiempo 4", value: randTime(), type: "time" }
                ]
            },
            {
                name: "Gasolina",
                value: randGas(),
                type: "currency"
            }
        ]
    };

    // 🎯 MATEMÁTICA VDT REAL: Calcular Costos y EBITDA
    function calculateVDT(node) {
        if (!node.children) return node.value;

        if (node.name === "EBITDA") {
            // EBITDA = Ingreso Base (Ej: $50,000) - Todos los costos
            let totalCostos = node.children.reduce((acc, child) => acc + calculateVDT(child), 0);
            node.value = 50000 - totalCostos; 
        } 
        else if (node.name.includes("Costo")) {
            // Costo = Suma de tiempos (minutos) * Tarifa de equipo (Ej: $450 por minuto)
            let totalMinutos = node.children.reduce((acc, child) => acc + calculateVDT(child), 0);
            node.value = totalMinutos * 450; 
        }
        return node.value;
    }
    // Ejecutamos la matemática para llenar los valores
    calculateVDT(data);

    // 📏 Dimensiones
    const width = 900;
    const height = 550;

    const svg = d3.select("#p3-vdt-canvas")
        .append("svg")
        .attr("width", "100%")
        .attr("height", "100%")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("preserveAspectRatio", "xMidYMid meet");

    const g = svg.append("g").attr("transform", "translate(120, 50)");

    const tree = d3.tree().size([height - 100, width - 350]);
    let root = d3.hierarchy(data);
    
    root.x0 = height / 2;
    root.y0 = 0;
    let i = 0;

    function update(source) {
        const treeData = tree(root);
        const nodes = treeData.descendants();
        const links = treeData.links();

        nodes.forEach(d => d.y = d.depth * 250);

        // --- LINKS ---
        const link = g.selectAll(".link").data(links, d => d.target.id);
        const linkEnter = link.enter()
            .insert("path", "g")
            .attr("class", "link")
            .attr("fill", "none")
            .attr("stroke", "#555")
            .attr("stroke-width", 2)
            .attr("d", d => {
                const o = { x: source.x0, y: source.y0 };
                return d3.linkHorizontal().x(d => d.y).y(d => d.x)({ source: o, target: o });
            });

        const linkUpdate = linkEnter.merge(link);
        linkUpdate.transition().duration(400)
            .attr("d", d3.linkHorizontal().x(d => d.y).y(d => d.x));

        link.exit().transition().duration(400)
            .attr("d", d => {
                const o = { x: source.x, y: source.y };
                return d3.linkHorizontal().x(d => d.y).y(d => d.x)({ source: o, target: o });
            }).remove();

        // --- NODES ---
        const node = g.selectAll(".node").data(nodes, d => d.id || (d.id = ++i));

        const nodeEnter = node.enter()
            .append("g")
            .attr("class", "node")
            .attr("transform", d => `translate(${source.y0},${source.x0})`)
            .on("click", (event, d) => toggle(d))
            .style("cursor", "pointer");

        // Caja más alta para que quepan 2 líneas de texto
        nodeEnter.append("rect")
            .attr("width", 160)
            .attr("height", 54) 
            .attr("x", -80)
            .attr("y", -27)
            .attr("rx", 8)
            .attr("fill", d => {
                if (d.depth === 0) return "#16a34a"; // EBITDA (Verde)
                if (d.depth === 1) return "#dc2626"; // Costos (Rojo)
                return "#2563eb"; // Tiempos (Azul)
            });

        // Agregando texto en 2 líneas usando <tspan>
        nodeEnter.append("text")
            .attr("text-anchor", "middle")
            .attr("fill", "#fff")
            .style("font-family", "sans-serif")
            .each(function(d) {
                const el = d3.select(this);
                
                // Línea 1: Nombre (Ej: "Tiempo 1" o "Costo Camiones")
                el.append("tspan")
                  .attr("x", 0)
                  .attr("dy", "-4") // Sube un poco
                  .style("font-size", "12px")
                  .text(d.data.name);
                  
                // Línea 2: Valor formateado
                el.append("tspan")
                  .attr("x", 0)
                  .attr("dy", "18") // Baja a la siguiente línea
                  .style("font-size", "15px")
                  .style("font-weight", "bold")
                  .text(() => {
                      if (d.data.type === "time") {
                          // Formato: 1,34 min
                          return d.data.value.toFixed(2).replace('.', ',') + " min";
                      } else {
                          // Formato: $45.000
                          return "$" + Math.floor(d.data.value).toLocaleString('es-CO');
                      }
                  });
            });

        const nodeUpdate = nodeEnter.merge(node);
        nodeUpdate.transition().duration(400)
            .attr("transform", d => `translate(${d.y},${d.x})`);

        const nodeExit = node.exit().transition().duration(400)
            .attr("transform", d => `translate(${source.y},${source.x})`)
            .remove();
        
        nodeExit.select("rect").attr("width", 0).attr("height", 0);
        nodeExit.select("text").style("fill-opacity", 0);

        nodes.forEach(d => {
            d.x0 = d.x;
            d.y0 = d.y;
        });
    }

    function toggle(d) {
        if (d.children) {
            d._children = d.children;
            d.children = null;
        } else {
            d.children = d._children;
            d._children = null;
        }
        update(d);
    }

    update(root);
}


// ==========================================
// LÓGICA DEL TORNADO CHART INTERACTIVO (PROYECTO 3)
// ==========================================

function initTornadoChart() {
    const containerNode = document.getElementById("tornado-chart");
    if (!containerNode) return;
    
    // Evitar dibujar el gráfico dos veces si ya existe
    if (containerNode.querySelector("svg")) return;

    const container = d3.select("#tornado-chart");
    container.selectAll("*").remove();

    const width = 900;
    const height = 500;
    const margin = { top: 20, right: 80, bottom: 40, left: 250 };

    // Hacemos el SVG responsivo con viewBox
    const svg = container.append("svg")
        .attr("width", "100%")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("preserveAspectRatio", "xMidYMid meet");

    const chart = svg.append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    // 🎯 Variables
    const variables = [
        "Loading time (shovel)",
        "Hauling time (truck)",
        "Dumping time",
        "Queue time at shovel",
        "Queue time at dump",
        "Spotting time",
        "Cycle time total",
        "Availability of trucks",
        "Availability of shovels",
        "Payload variability",
        "Distance to dump",
        "Fuel efficiency impact",
        "Maintenance downtime",
        "Operator efficiency",
        "Idle time"
    ];

    // 🎲 Generar datos aleatorios
    const data = variables.map(v => {
        const min = +(Math.random() * 2 + 0.5).toFixed(2);
        const max = +(Math.random() * 2 + 2.5).toFixed(2);
        return { name: v, min, max };
    });

    const baseline = 2.5;

    const y = d3.scaleBand()
        .domain(data.map(d => d.name))
        .range([0, height - margin.top - margin.bottom])
        .padding(0.2);

    const x = d3.scaleLinear()
        .domain([0.5, 5])
        .range([0, width - margin.left - margin.right]);

    // 🎯 Línea base
    chart.append("line")
        .attr("x1", x(baseline))
        .attr("x2", x(baseline))
        .attr("y1", 0)
        .attr("y2", height - margin.top - margin.bottom)
        .attr("stroke", "#fff")
        .attr("stroke-width", 2);

    // 🟩 MIN (izquierda)
    chart.selectAll(".bar-min")
        .data(data)
        .enter()
        .append("rect")
        .attr("class", "bar-min")
        .attr("y", d => y(d.name))
        .attr("x", d => x(d.min))
        .attr("width", d => x(baseline) - x(d.min))
        .attr("height", y.bandwidth())
        .attr("fill", "#4ade80")
        .on("mousemove", showTooltip)
        .on("mouseout", hideTooltip);

    // 🟦 MAX (derecha)
    chart.selectAll(".bar-max")
        .data(data)
        .enter()
        .append("rect")
        .attr("class", "bar-max")
        .attr("y", d => y(d.name))
        .attr("x", x(baseline))
        .attr("width", d => x(d.max) - x(baseline))
        .attr("height", y.bandwidth())
        .attr("fill", "#1e3a8a")
        .on("mousemove", showTooltip)
        .on("mouseout", hideTooltip);

    // 📊 Eje Y
    chart.append("g")
        .call(d3.axisLeft(y))
        .selectAll("text")
        .style("fill", "#cbd5f5")
        .style("font-size", "13px")
        .style("font-family", "sans-serif");

    // 📊 Eje X
    chart.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).ticks(10).tickFormat(d => `$${d}/ton`))
        .selectAll("text")
        .style("fill", "#94a3b8")
        .style("font-family", "sans-serif");

    // Quitar la línea del eje X y Y para un diseño más limpio
    chart.selectAll(".domain").remove();

    // 🧠 TOOLTIP
    const tooltip = d3.select("body")
        .append("div")
        .style("position", "absolute")
        .style("background", "#020617")
        .style("color", "#fff")
        .style("padding", "8px 12px")
        .style("border-radius", "6px")
        .style("font-size", "13px")
        .style("font-family", "sans-serif")
        .style("border", "1px solid #334155")
        .style("pointer-events", "none")
        .style("opacity", 0)
        .style("z-index", 1000)
        .style("box-shadow", "0 4px 6px rgba(0,0,0,0.5)");

    function showTooltip(event, d) {
        tooltip.style("opacity", 1)
            .html(`
                <strong style="color: #4ade80;">${d.name}</strong><br>
                Min: $${d.min}/ton<br>
                Max: $${d.max}/ton
            `)
            .style("left", (event.pageX + 15) + "px")
            .style("top", (event.pageY - 25) + "px");
        
        // Efecto hover (ilumina la barra)
        d3.select(this).attr("opacity", 0.7);
    }

    function hideTooltip() {
        tooltip.style("opacity", 0);
        d3.select(this).attr("opacity", 1);
    }
}

// ==========================================
// OBSERVADOR UNIFICADO (ÚNICO)
// ==========================================
const observer = new MutationObserver(() => {
    // Si existe el canvas del Árbol, lo dibuja
    if (document.getElementById("p3-vdt-canvas")) {
        renderVDT();
    }
    // Si existe el canvas del Tornado, lo dibuja
    if (document.getElementById("tornado-chart")) {
        initTornadoChart();
    }
});
observer.observe(document.body, { childList: true, subtree: true });

}

// ==========================================
// LÓGICA DEL LIGHTBOX (IMÁGENES EN GRANDE)
// ==========================================
document.addEventListener('click', function(e) {
    // Usamos querySelector para leer las clases que ya tienes en tu CSS
    const modal = document.querySelector('.lightbox-modal');
    const modalImg = document.querySelector('.lightbox-content');
    
    // 1. Si hace clic en una imagen con la clase 'clickable-img', abre el modal
    if (e.target && e.target.classList.contains('clickable-img')) {
        if (modal && modalImg) {
            modalImg.src = e.target.src; // Copia la ruta de la imagen tocada
            modal.classList.add('show');
        }
    }
    
    // 2. Si hace clic en la "X" o en cualquier parte del fondo borroso, lo cierra
    if (e.target && (e.target.classList.contains('lightbox-modal') || e.target.classList.contains('lightbox-close'))) {
        if (modal) {
            modal.classList.remove('show');
        }
    }
}); // <-- ¡Paréntesis y punto y coma cerrados correctamente!







