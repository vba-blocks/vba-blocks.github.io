---
title: Blog
layout: default
---

{% for post in site.posts %}
<section>
<h3><a href="{{ post.url }}">{{ post.title }}</a></h3>
<p class="faded mb1">
  {% if post.author %}{{ post.author }}{% else %}Tim Hall{% endif %} /
  {{ post.date | date: "%b %-d, %Y" }}
  {% for tag in post.tags %}
    / <!--<a href="{{ site.url }}/blog?tag={{ tag }}">-->{{ tag }}<!--</a>-->
  {% endfor %}
</p>
<p>{{ post.description}}</p>
</section>
{% else %}

Upcoming posts on vba-blocks:

- Building vba-blocks
- Working with vba-blocks
- Manifest deep dive
- Lots more!

Check back soon!

{% endfor %}