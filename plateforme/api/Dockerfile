# Construction puis image minimale : le binaire Go est
# statique, il n'a besoin ni de runtime ni de bibliotheque.

FROM golang:1.25-alpine AS construction
WORKDIR /src

# Les dependances d'abord : cette couche se recycle tant que
# go.mod et go.sum ne bougent pas.
COPY go.mod go.sum ./
RUN go mod download

COPY . .
RUN CGO_ENABLED=0 go build -ldflags="-s -w" -o /bin/api .

FROM gcr.io/distroless/static-debian12:nonroot
COPY --from=construction /bin/api /api
EXPOSE 8080
USER nonroot:nonroot
ENTRYPOINT ["/api"]
