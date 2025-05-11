Run the following command to produce the image:

```bash
docker build -t grade-extractor .
```

Assuming the PDF lives in the current local folder, run the container and mount in the PDF running:

```bash
docker run --rm -it \
  -v "$(pwd)":/data \
  -w /data \
  grade-extractor
```

Inside the container you’ll find your PDF in `/data`. Run the script with:

```bash
python esfera-acta-extractor.py
```
