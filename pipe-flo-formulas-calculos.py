import math

def calcular_ft(d, e):
    """
    Calcula el factor de fricción turbulenta (fT) usando la ecuación de Nikuradse (Ecuación 9).
    d: Diámetro de la tubería.
    e: Rugosidad superficial del material de la tubería.
    Ambas unidades deben ser consistentes (ej. mm).
    """
    if d <= 0 or e <= 0:
        raise ValueError("El diámetro y la rugosidad deben ser mayores que cero.")
    
    # fT = 8 * [2.45 * ln(3.707 * d / e)]^-2
    termino_ln = math.log(3.707 * d / e)
    ft = 8 * (2.45 * termino_ln) ** -2
    return ft

def calcular_k_desde_dp(dp, d, rho, w):
    """
    Calcula el valor K necesario para lograr una presión diferencial específica (Ecuación 10).
    dp: Cambio de presión (bares)
    d: Diámetro de la tubería (mm)
    rho: Densidad del fluido (kg/m3)
    w: Tasa de flujo másico (kg/hr)
    """
    # K = 1.59923 * dP * d^4 * rho / W^2
    k = (1.59923 * dp * (d ** 4) * rho) / (w ** 2)
    return k

def calcular_k_desde_cv(d_pulgadas, cv):
    """
    Convierte el coeficiente de flujo Cv al coeficiente de resistencia K (Ecuación 11).
    d_pulgadas: Diámetro interior de la tubería (pulgadas)
    cv: Coeficiente de flujo (gpm de agua a 60 F para caída de 1 psi)
    """
    if cv == 0:
        raise ValueError("Cv no puede ser cero.")
    # K = 891 * d^4 / Cv^2
    k = (891 * (d_pulgadas ** 4)) / (cv ** 2)
    return k

def formula_1_contraccion(theta_grados, d_menor, d_mayor):
    """
    Formula 1: Calcula el valor K para contracciones de tubería (Ecuaciones 12a y 12b).
    theta_grados: Ángulo de aproximación en grados.
    d_menor: Diámetro menor (d1).
    d_mayor: Diámetro mayor (d2).
    """
    beta = d_menor / d_mayor
    theta_rad = math.radians(theta_grados)
    
    if theta_grados <= 45:
        # K1 = 0.8 * sin(theta/2) * (1 - beta^2)
        k = 0.8 * math.sin(theta_rad / 2) * (1 - beta ** 2)
    elif 45 < theta_grados <= 180:
        # K1 = 0.5 * (1 - beta^2) * (sin(theta/2))^(1/2)
        k = 0.5 * (1 - beta ** 2) * math.sqrt(math.sin(theta_rad / 2))
    else:
        raise ValueError("El ángulo theta debe estar entre 0 y 180 grados.")
    
    return k

def formula_3_ampliacion(theta_grados, d_menor, d_mayor):
    """
    Formula 3: Calcula el valor K para ampliaciones de tubería (Ecuaciones 13a y 13b).
    theta_grados: Ángulo de aproximación en grados.
    d_menor: Diámetro menor (d1).
    d_mayor: Diámetro mayor (d2).
    """
    beta = d_menor / d_mayor
    theta_rad = math.radians(theta_grados)
    
    if theta_grados <= 45:
        # K1 = 2.6 * sin(theta/2) * (1 - beta^2)^2
        k = 2.6 * math.sin(theta_rad / 2) * ((1 - beta ** 2) ** 2)
    elif 45 < theta_grados <= 180:
        # K1 = (1 - beta^2)^2
        k = (1 - beta ** 2) ** 2
    else:
        raise ValueError("El ángulo theta debe estar entre 0 y 180 grados.")
    
    return k

def formula_5_valvula_asiento_reducido_gradual(theta_grados, d_valvula, d_tuberia, k1_asiento_reducido):
    """
    Formula 5: Válvulas de asiento reducido con cambio gradual de diámetro (Ecuación 14).
    Ejemplos: Válvulas de bola y compuerta.
    theta_grados: Ángulo de aproximación en grados.
    d_valvula: Diámetro del asiento de la válvula.
    d_tuberia: Diámetro de la tubería.
    k1_asiento_reducido: Valor K propio del asiento reducido de la válvula.
    """
    beta = d_valvula / d_tuberia
    
    k_reducer = formula_1_contraccion(theta_grados, d_valvula, d_tuberia)
    k_enlarger = formula_3_ampliacion(theta_grados, d_valvula, d_tuberia)
    
    # K2 = K_Reducer + K1 / beta^4 + K_Enlarger
    k2 = k_reducer + (k1_asiento_reducido / (beta ** 4)) + k_enlarger
    return k2

def formula_7_valvula_asiento_reducido_abrupto(d_valvula, d_tuberia, k1_asiento_reducido):
    """
    Formula 7: Válvulas de asiento reducido con cambio abrupto de diámetro.
    Ejemplos: Válvulas de globo, ángulo, retención de elevación.
    Es la Fórmula 5 con theta = 180 grados.
    """
    return formula_5_valvula_asiento_reducido_gradual(180, d_valvula, d_tuberia, k1_asiento_reducido)

def formula_8_codos_y_curvas(n_codos_90, fT, r_d, k_un_codo_90):
    """
    Formula 8: Calcula el valor K para codos y curvas.
    n_codos_90: Número de curvas de 90 grados (n).
    fT: Factor de fricción turbulenta.
    r_d: Relación radio a diámetro (r/d).
    k_un_codo_90: Coeficiente de resistencia para una curva de 90 grados (K).
    """
    # Kb = (n - 1) * (0.25 * pi * fT * (r/d) + 0.5 * K) + K
    kb = (n_codos_90 - 1) * (0.25 * math.pi * fT * r_d + 0.5 * k_un_codo_90) + k_un_codo_90
    return kb

def formulas_9_y_10_ld(fT, relacion_L_D):
    """
    Formulas 9 y 10: Válvulas con L/D que varía con el diámetro, 
    y válvulas/accesorios de asiento completo.
    Ejemplos: Válvulas mariposa, retención de disco basculante, tapón, pie, tes, etc.
    fT: Factor de fricción turbulenta.
    relacion_L_D: Coeficiente L/D de la válvula o accesorio.
    """
    # K = fT * (L/D)
    return fT * relacion_L_D

def formula_11_k_fijo(k_fijo):
    """
    Formula 11: Accesorios con valor K fijo (ej. entradas y salidas de tubería).
    Retorna el mismo valor K. Se incluye para completitud arquitectónica del código.
    """
    return k_fijo

# === Ejemplo de Uso Rápido ===
if __name__ == "__main__":
    # Datos de prueba arbitrarios
    diametro_tuberia = 100 # mm
    rugosidad = 0.045 # mm (Acero comercial típico)
    
    f_turbulento = calcular_ft(diametro_tuberia, rugosidad)
    print(f"Factor de fricción turbulenta (fT): {f_turbulento:.4f}")
    
    k_contraccion = formula_1_contraccion(30, 50, 100)
    print(f"K Contracción (theta=30, d1=50, d2=100): {k_contraccion:.4f}")
    
    k_valvula_globo_abrupta = formula_7_valvula_asiento_reducido_abrupto(80, 100, 2.5)
    print(f"K Válvula Asiento Reducido Abrupto: {k_valvula_globo_abrupta:.4f}")