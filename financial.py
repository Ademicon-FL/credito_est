"""
financial.py
Funções financeiras puras (sem numpy_financial).
Usa scipy.optimize.brentq para IRR - mais robusto que Newton-Raphson.
"""
import numpy as np
try:
    from scipy.optimize import brentq
    SCIPY_OK = True
except ImportError:
    SCIPY_OK = False


def irr(cashflows: list) -> float:
    """
    Calcula a TIR (taxa interna de retorno) dos fluxos de caixa.
    Retorna a taxa mensal. Retorna 0.01 se não convergir.
    """
    cfs = np.array(cashflows, dtype=float)
    
    def npv(r):
        t = np.arange(len(cfs))
        return np.sum(cfs / (1.0 + r) ** t)
    
    # Tenta brentq (robusto) primeiro
    if SCIPY_OK:
        try:
            for lo, hi in [(-0.5, 5.0), (-0.9, 50.0)]:
                try:
                    if npv(lo) * npv(hi) < 0:
                        return float(brentq(npv, lo, hi, xtol=1e-10, maxiter=500))
                except Exception:
                    pass
        except Exception:
            pass
    
    # Fallback: Newton-Raphson
    r = 0.01
    for _ in range(500):
        t = np.arange(len(cfs))
        f  = np.sum(cfs / (1 + r) ** t)
        df = np.sum(-t * cfs / (1 + r) ** (t + 1))
        if abs(df) < 1e-14:
            break
        r_new = r - f / df
        if r_new < -0.999:
            r_new = -0.5
        if abs(r_new - r) < 1e-10:
            r = r_new
            break
        r = r_new
    
    return float(r) if not np.isnan(r) and -1 < r < 100 else 0.01
