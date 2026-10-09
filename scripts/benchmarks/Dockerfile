FROM python@sha256:2325bb286ec344af3e5898cc224b5844e2707ac6e26b1632516fd3edc84a5e26
ARG SHEETOPS=0
RUN pip install --no-cache-dir openpyxl==3.1.5 et-xmlfile==2.0.0 \
    && if [ "$SHEETOPS" = "1" ]; then pip install --no-cache-dir --no-deps zalos-sheetops==0.4.2; fi
COPY xlsx-scripts /opt/skills/xlsx/scripts
WORKDIR /workspace
CMD ["sleep", "infinity"]
