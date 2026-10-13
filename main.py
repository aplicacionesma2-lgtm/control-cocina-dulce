# ---------------------------------------------------------
# PESTAÑA 2: MODIFICAR / ELIMINAR REGISTROS
# ---------------------------------------------------------
with tab2:
    st.subheader("✏️ Editar o Eliminar un Registro Existente")
    if sheet is not None:
        try:
            raw_data = sheet.get_all_values()
            if len(raw_data) > 1:
                headers = raw_data[0]
                rows = raw_data[1:]

                df_edit = pd.DataFrame(rows, columns=headers)
                df_edit["Fila GS"] = list(range(2, len(rows) + 2))

                opciones_registro = [
                    f"Fila {r['Fila GS']} | {r.get('PRODUCTO', '')} | Lote: {r.get('LOTE', '')} | FP: {r.get('F.P', '')}"
                    for _, r in df_edit.iterrows()
                ]

                registro_sel = st.selectbox(
                    "Seleccione el registro que desea modificar o eliminar:",
                    options=opciones_registro,
                )

                idx_seleccionado = opciones_registro.index(registro_sel)
                fila_gs = df_edit.iloc[idx_seleccionado]["Fila GS"]
                datos_fila = df_edit.iloc[idx_seleccionado].to_dict()

                col_mod, col_del = st.columns(2)

                # --- ACCIÓN 1: MODIFICAR ---
                with col_mod:
                    with st.expander(f"📝 Editar datos de la Fila {fila_gs}", expanded=True):
                        with st.form(f"form_editar_{fila_gs}"):
                            e_prod = st.text_input("PRODUCTO", value=datos_fila.get("PRODUCTO", ""))
                            e_fp = st.text_input("F.P", value=datos_fila.get("F.P", ""))
                            e_lote = st.text_input("LOTE", value=datos_fila.get("LOTE", ""))
                            e_batch = st.text_input("BATCH", value=datos_fila.get("BATCH", ""))
                            e_equipo = st.text_input("EQUIPO UTILIZADO", value=datos_fila.get("EQUIPO UTILIZADO", ""))
                            e_brix = st.text_input("BRIX (°Bx)", value=datos_fila.get("BRIX (°Bx)", ""))
                            e_h_ini = st.text_input("HORA INICIO", value=datos_fila.get("HORA INICIO", ""))
                            e_temp = st.text_input("TEMPERATURA EQUIPO (°C)", value=datos_fila.get("TEMPERATURA EQUIPO (°C)", datos_fila.get("TEMPERATURA  EQUIPO (°C)", "")))
                            
                            # Carga flexible para Tiempo y Velocidad
                            e_tiempo = st.text_input(
                                "TIEMPO", 
                                value=datos_fila.get("TIEMPO", datos_fila.get("TIEMPO (COCCIÓN/BATIDO/ HORNEADO)", ""))
                            )
                            e_vel = st.text_input(
                                "VELOCIDAD AGITADOR", 
                                value=datos_fila.get(
                                    "VELOCIDAD DEL AGITADOR (hz)/ BATIDORA/OTROS", 
                                    datos_fila.get("VELOCIDAD AGITADOR/ BATIDORA", datos_fila.get("VELOCIDAD AGITADOR", ""))
                                )
                            )
                            
                            e_h_term = st.text_input("HORA TÉRMINO", value=datos_fila.get("HORA TÉRMINO", ""))
                            e_resp = st.text_input("RESPONSABLE", value=datos_fila.get("RESPONSABLE", ""))
                            e_obs = st.text_area("OBSERVACIÓN", value=datos_fila.get("OBSERVACIÓN", ""))
                            e_cond_area = st.text_input("CONDICIONES ÁREA TRABAJO", value=datos_fila.get("CONDICIONES DEL AREA DE TRABAJO", datos_fila.get("CONDICIONES DEL ÁREA DE TRABAJO", "")))

                            btn_actualizar = st.form_submit_button("💾 Actualizar Fila en Google Sheets")

                        if btn_actualizar:
                            valores_actualizados = [
                                e_prod, e_fp, e_lote, e_batch, e_equipo,
                                e_brix, e_h_ini, e_temp, e_tiempo, e_vel,
                                e_h_term, e_resp, e_obs, e_cond_area
                            ]
                            sheet.update(
                                range_name=f"A{fila_gs}:N{fila_gs}",
                                values=[valores_actualizados],
                                value_input_option="USER_ENTERED"
                            )
                            st.success(f"✅ ¡Fila {fila_gs} actualizada correctamente!")
                            st.cache_data.clear()
                            st.rerun()

                # --- ACCIÓN 2: ELIMINAR ---
                with col_del:
                    with st.expander(f"🗑️ Eliminar Fila {fila_gs}", expanded=True):
                        st.warning(f"⚠️ ¿Estás seguro de que deseas eliminar permanentemente el registro de la fila {fila_gs}?")
                        btn_eliminar = st.button(f"❌ Sí, Eliminar Fila {fila_gs}", type="primary", key=f"btn_del_{fila_gs}")

                        if btn_eliminar:
                            sheet.delete_rows(int(fila_gs))
                            st.success(f"🗑️ ¡Fila {fila_gs} eliminada correctamente!")
                            st.cache_data.clear()
                            st.rerun()

            else:
                st.info("No hay registros suficientes para modificar.")
        except Exception as e:
            st.error(f"Error al cargar registros para edición: {e}")
