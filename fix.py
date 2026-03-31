content = open('app.py', encoding='utf-8').read()

old = '''    st.dataframe(df_show.style
        .background_gradient(cmap="YlOrRd", subset=["LeadScore","OralBioScore"])
        .background_gradient(cmap="Reds",   subset=["PromiscuityRisk","SA_Score","CYP_Hits"])
        .background_gradient(cmap="Blues",  subset=["Sim"])
        .background_gradient(cmap="Greens", subset=["QED"]),
        use_container_width=True, height=min(80+34*total,320))'''

new = '''    st.dataframe(df_show, use_container_width=True, height=min(80+34*total,320))'''

if old in content:
    content = content.replace(old, new)
    print('FIXED')
else:
    print('NOT FOUND - printing exact lines around 1912')
    lines = content.split('\n')
    for i in range(1905, 1920):
        print(f"{i+1}: {repr(lines[i])}")

open('app.py', 'w', encoding='utf-8').write(content)