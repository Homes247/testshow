import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { environment } from '../../environments/environment';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class GiphyService {
  private apiKey = (environment as any).giphyApiKey || 'Gc7131jiJuvI7IdN0HZ1D7nh0ow5BU6g';
  private baseUrl = 'https://api.giphy.com/v1';

  constructor(private http: HttpClient) {}

  searchGifs(query: string, limit: number = 20): Observable<any> {
    return this.http.get(`${this.baseUrl}/gifs/search?api_key=${this.apiKey}&q=${encodeURIComponent(query)}&limit=${limit}&rating=g`);
  }

  getTrendingGifs(limit: number = 20): Observable<any> {
    return this.http.get(`${this.baseUrl}/gifs/trending?api_key=${this.apiKey}&limit=${limit}&rating=g`);
  }

  searchStickers(query: string, limit: number = 20): Observable<any> {
    return this.http.get(`${this.baseUrl}/stickers/search?api_key=${this.apiKey}&q=${encodeURIComponent(query)}&limit=${limit}&rating=g`);
  }

  getTrendingStickers(limit: number = 20): Observable<any> {
    return this.http.get(`${this.baseUrl}/stickers/trending?api_key=${this.apiKey}&limit=${limit}&rating=g`);
  }
}
