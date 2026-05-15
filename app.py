# =====================================================
# TAB 3 - PROYECTOS COMPLETADOS
# =====================================================

with tab3:
    st.markdown(
        '<div class="section-title">Proyectos completados</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="note">Aquí aparecen únicamente los proyectos con estado Completo o Finalizado.</div>',
        unsafe_allow_html=True
    )

    df_completados = df_f[df_f["Es completo"]].copy()

    df_completados = df_completados.sort_values(
        by=["Sin fecha", "Fecha entrega"],
        ascending=[True, False]
    )

    if df_completados.empty:
        st.info("No hay proyectos completados con los filtros seleccionados.")

    else:
        total_completados = len(df_completados)
        avance_promedio_completados = round(df_completados["Avance"].mean(), 1)

        c1, c2 = st.columns(2)

        with c1:
            st.metric("Total completados", int(total_completados))

        with c2:
            st.metric("Avance promedio", f"{avance_promedio_completados}%")

        vista_completados = st.radio(
            "Vista de completados",
            ["Lista general", "Agrupado por tipo", "Tabla"],
            horizontal=True,
            key="vista_completados"
        )

        if vista_completados == "Lista general":
            for _, row in df_completados.iterrows():
                render_card(row)

        elif vista_completados == "Agrupado por tipo":
            tipos = list(df_completados["Tipo"].dropna().unique())

            orden_preferido = [
                "Estudio Recurrente",
                "Estudios Recurrentes",
                "Iniciativa",
                "Iniciativas",
                "Solicitud Interna",
                "Solicitudes Internas",
                "Tendencias"
            ]

            tipos_ordenados = []

            for tipo_preferido in orden_preferido:
                for tipo_real in tipos:
                    if str(tipo_real).strip().lower() == tipo_preferido.lower():
                        if tipo_real not in tipos_ordenados:
                            tipos_ordenados.append(tipo_real)

            for tipo_real in tipos:
                if tipo_real not in tipos_ordenados:
                    tipos_ordenados.append(tipo_real)

            num_cols = min(len(tipos_ordenados), 4)

            if num_cols == 0:
                st.warning("No hay proyectos completados para mostrar.")
            else:
                cols = st.columns(num_cols)

                for i, tipo in enumerate(tipos_ordenados):
                    with cols[i % num_cols]:
                        st.subheader(tipo)

                        subset = df_completados[df_completados["Tipo"] == tipo].sort_values(
                            by=["Sin fecha", "Fecha entrega"],
                            ascending=[True, False]
                        )

                        for _, row in subset.iterrows():
                            render_card(row)

        else:
            tabla_completados = df_completados[
                [
                    "Tipo",
                    "Proyecto",
                    "Mini proyecto",
                    "Descripción",
                    "Encargado",
                    "Equipo",
                    "Estado",
                    "Avance",
                    "Fecha texto"
                ]
            ].copy()

            tabla_completados = tabla_completados.rename(
                columns={
                    "Avance": "% completado",
                    "Fecha texto": "Fecha entrega"
                }
            )

            st.dataframe(
                tabla_completados,
                use_container_width=True,
                hide_index=True
            )
