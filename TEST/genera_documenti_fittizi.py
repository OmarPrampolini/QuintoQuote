from __future__ import annotations

import json
from pathlib import Path

import preventivo_generator_v2 as qq


ROOT = Path(__file__).resolve().parents[1]
TEST_DIR = ROOT / "TEST"
INPUT_DIR = TEST_DIR / "input"
VALUES_DIR = TEST_DIR / "values"
OUTPUT_DIR = TEST_DIR / "output"


BASE_CASE_DATA = {
    "full_name": "Mario Rossi",
    "birth_place": "Roma",
    "birth_date": "15/05/1980",
    "birth_province_name": "Roma",
    "birth_province_code": "RM",
    "tax_code": "RSSMRA80E15H501Z",
    "payroll_number": "X-80226489",
    "residence_city": "Roma",
    "residence_province_name": "Roma",
    "residence_province_code": "RM",
    "residence_cap": "00184",
    "residence_street": "Via Nazionale",
    "residence_number": "10",
    "phone": "3331234567",
    "fax": "0612345678",
    "email": "mario.rossi.demo@example.it",
    "service_office": "Ministero dell'Economia e delle Finanze",
    "employer_entity": "Ragioneria Territoriale dello Stato di Roma",
    "lender_name": "Finanziaria Demo S.p.A.",
    "contract_number": "QQ-TEST-2026-0001",
    "contract_date": "16/04/2026",
    "loan_type": "Delegazione di Pagamento",
    "iban": "IT60X0542811101000000123456",
    "loan_amount": "35.674,87",
    "net_disbursed": "35.674,87",
    "total_ceded": "45.120,00",
    "fees_total": "0,00",
    "interest_total": "9.445,13",
    "monthly_installment": "376,00",
    "salary_fifth": "400,00",
    "installment_count": "120",
    "tan": "4,86",
    "taeg": "4,98",
    "teg": "4,98",
    "insurance": "Polizza demo credito e morte n. 2026-001",
    "other_financing_lender": "Finanziaria Estinzione Demo S.p.A.",
    "other_financing_installment": "220,00",
    "other_financing_expiry": "31/12/2030",
}


MODULES = (
    (
        qq.ALLEGATO_E_SPEC,
        "allegato_e_fittizio.pdf",
        {
            "destinatario_riga_1": "Direzione dei Servizi del Tesoro",
            "destinatario_riga_2": "Ufficio Stipendi Centrali",
            "destinatario_riga_3": "Roma",
            "importo_finanziamento_lettere": "trentacinquemilaseicentosettantaquattro/87",
            "importo_globale_ceduto_lettere": "quarantacinquemilacentoventi/00",
            "luogo_timbro_istituto": "Roma",
            "data_timbro_istituto": "16/04/2026",
            "documento_identificazione": "Carta di identita n. AA1234567",
            "luogo_autentica": "Roma",
            "data_autentica": "16/04/2026",
        },
    ),
    (
        qq.ALLEGATO_C_SPEC,
        "allegato_c_fittizio.pdf",
        {
            "data_modulo": "16/04/2026",
        },
    ),
    (
        qq.FRONTESPIZIO_BANCHE_SPEC,
        "frontespizio_banche_fittizio.pdf",
        {
            "cip": "1234",
            "nome_sede": "Roma",
            "doc_allegato_1": "contratto.pdf",
            "doc_allegato_2": "allegato_E.pdf",
            "doc_allegato_3": "documento_identificativo.pdf",
            "doc_allegato_4": "polizza_assicurativa.pdf",
            "doc_allegato_5": "consenso_noipa.pdf",
            "eventuali_comnunicazioni_alla_rts": "Pratica demo generata per test funzionale interno QuintoQuote.",
            "codice_delegato": "D1234",
            "denominazione_delegato": "Agenzia Demo Roma Centro",
            "email_delegato": "agenzia.demo@example.it",
            "telefono_delegato": "0611122233",
            "tipologia_estinzione_2": "Delega-Piccolo Prestito-Prestito Doppio",
            "societa_estinzione_2": "Prestito Demo 2 S.r.l.",
            "rata_estinzione_2": "150,00",
            "tipologia_estinzione_3": "Pignoramento-Alimenti-Recupero Obbligatorio",
            "societa_estinzione_3": "Recupero Demo 3",
            "rata_estinzione_3": "80,00",
        },
    ),
    (
        qq.FRONTESPIZIO_INTEGRATIVO_SPEC,
        "frontespizio_integrativo_fittizio.pdf",
        {
            "protocollo": "PROT-TEST-001",
            "data_protocollo": "17/04/2026",
            "osservazioni_rts": "Documentazione integrativa fittizia prodotta per smoke test del compilatore XFA.",
            "cip": "1234",
            "nome_sede": "Roma",
            "documento_integrativo_1": "Conteggio_Estintivo",
            "documento_integrativo_2": "Contabile_Bonifico",
            "documento_integrativo_3": "Dichiarazione_per_estinzione",
            "documento_integrativo_4": "1",
            "Societa_erogatrice": "Agenzia Demo Roma Centro",
            "denominazione_delegato": "Agenzia Demo Roma Centro",
            "codice_delegato": "D1234",
        },
    ),
)


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True), encoding="utf-8")


def main() -> None:
    for folder in (INPUT_DIR, VALUES_DIR, OUTPUT_DIR):
        folder.mkdir(parents=True, exist_ok=True)

    _write_json(INPUT_DIR / "pratica_fittizia.json", BASE_CASE_DATA)

    generated_files: list[str] = []
    for spec, output_name, overrides in MODULES:
        raw_values = qq.build_prefill_for_template(spec.key, BASE_CASE_DATA)
        raw_values.update(overrides)
        clean_values = qq.sanitize_pdf_form_payload(spec, raw_values)
        output_path = OUTPUT_DIR / output_name
        qq.render_pdf_template(spec, clean_values, output_path)
        _write_json(VALUES_DIR / f"{spec.slug}.json", clean_values)
        generated_files.append(output_name)

    summary_lines = [
        "# RIEPILOGO TEST",
        "",
        "PDF generati con dati totalmente fittizi:",
        "",
    ]
    for filename in generated_files:
        summary_lines.append(f"- `{filename}`")
    summary_lines.extend(
        [
            "",
            "File input base:",
            f"- `input/pratica_fittizia.json`",
            "",
            "Valori usati per ciascun modulo:",
            "- `values/allegato-e.json`",
            "- `values/allegato-c.json`",
            "- `values/frontespizio-banche.json`",
            "- `values/frontespizio-integrativo.json`",
        ]
    )
    (OUTPUT_DIR / "RIEPILOGO.md").write_text("\n".join(summary_lines) + "\n", encoding="utf-8")

    print("Test completato.")
    for filename in generated_files:
        print(f"- {OUTPUT_DIR / filename}")


if __name__ == "__main__":
    main()
